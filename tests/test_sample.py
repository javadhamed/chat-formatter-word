"""Checks the pipeline model against the real sample document.

Run:  python3 tests/test_sample.py [path/to/docx-or-txt]
"""

import io
import os
import re
import sys
import time
import zipfile
from xml.etree import ElementTree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pipeline_ref as P

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DEFAULT_SAMPLE = os.path.expanduser("~/Downloads/لیست کامل و به.docx")


def read_docx_paragraphs(path):
    """Extract paragraph text straight from word/document.xml."""
    with zipfile.ZipFile(path) as z:
        root = ElementTree.fromstring(z.read("word/document.xml"))
    out = []
    for p in root.iter(W + "p"):
        buf = []
        for node in p.iter():
            if node.tag == W + "t":
                buf.append(node.text or "")
            elif node.tag == W + "tab":
                buf.append("\t")
        out.append("".join(buf))
    return out


def body_lines(doc):
    """Paragraphs that are not table placeholders, in order."""
    return [l for l in doc.lines if not l.startswith("\x00T")]


def all_lines(doc):
    """Everything the document now contains, table cells included.

    The no-data-loss check has to look inside the tables, otherwise a
    formatter that silently ate the tables would look lossless.
    """
    out = []
    for line in doc.lines:
        m = re.match(r"^\x00T(\d+)\x00$", line)
        if m:
            out.extend("|".join(row) for row in doc.tables[int(m.group(1))])
        else:
            out.append(line)
    return out


def content_words(lines):
    """Normalised text used for the no-data-loss comparison.

    Markdown punctuation that the formatter is supposed to strip is removed
    on both sides, so the comparison measures real content.
    """
    joined = "\n".join(lines)
    joined = re.sub(r"[`*_~>#|=\-]", "", joined)
    joined = re.sub(r"\s+", "", joined)
    return joined


def check(name, ok, detail=""):
    print(("  PASS  " if ok else "  FAIL  ") + name + (("  -- " + detail) if detail else ""))
    return bool(ok)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SAMPLE
    if not os.path.exists(path):
        print("sample not found: " + path)
        return 1

    lines = read_docx_paragraphs(path)
    ok = True

    print("sample: %s" % path)
    print("  paragraphs=%d  chars=%d" % (len(lines), sum(len(l) for l in lines)))

    # -- classification ---------------------------------------------------
    t0 = time.time()
    blocks = P.scan(lines)
    scan_s = time.time() - t0
    kinds = {}
    for b in blocks:
        kinds[b.kind] = kinds.get(b.kind, 0) + 1
    print("  blocks: %s" % kinds)

    ok &= check("scan is linear in paragraphs",
                len(blocks) <= len(lines),
                "%d blocks for %d paragraphs" % (len(blocks), len(lines)))

    # -- the actual run ---------------------------------------------------
    t0 = time.time()
    doc = P.run(lines)
    run_s = time.time() - t0
    print("  ops=%d  ratio=%.2f/para  (%.0f ms)" %
          (doc.ops, doc.ops / max(1, len(lines)), run_s * 1000))

    ok &= check("work is linear, not quadratic",
                doc.ops <= 6 * len(lines),
                "%d ops for %d paragraphs (quadratic would be ~%d)"
                % (doc.ops, len(lines), len(lines) ** 2 // 10))

    # -- no data loss -----------------------------------------------------
    # Fence markers and their language tags are syntax, not content, so
    # they are excluded from the source side of the comparison.
    payload = [l for l in lines if P.fence_info(l) is None]
    before = content_words(payload)
    after = content_words(all_lines(doc))
    ok &= check("no text is lost", before == after,
                "before=%d chars after=%d chars" % (len(before), len(after)))

    # -- fenced TSV must become real tables -------------------------------
    ok &= check("31 TSV fences became tables", len(doc.tables) == 31,
                "got %d tables" % len(doc.tables))

    grid_ok = all(g and len(g) >= 2 for g in doc.tables)
    ok &= check("every table has >= 2 columns", grid_ok)

    # -- code fences must NOT have been eaten by the TSV path -------------
    leftover = [l for l in doc.lines if "```" in l]
    ok &= check("no fence markers survive", not leftover,
                "%d left" % len(leftover))

    # -- headings ---------------------------------------------------------
    heads = [l for l in body_lines(doc) if P.ATX.match(l.strip())]
    ok &= check("ATX markers consumed", not heads, "%d left" % len(heads))

    out = "\n".join(body_lines(doc))
    ok &= check("no setext underlines remain",
                not re.search(r"^={2,}$", out, re.M))

    # -- lists ------------------------------------------------------------
    # The sample is a specification with no top-level markdown lists, so
    # these are asserted on a synthetic fixture instead (test_cases.py).
    src_bullets = sum(1 for l in payload if P.BULLET.match(l.strip()))
    src_numbered = sum(1 for l in payload if P.NUMBERED.match(l.strip()))
    if src_bullets or src_numbered:
        ok &= check("bullets converted",
                    out.count("\u2022\t") >= src_bullets,
                    "%d of %d" % (out.count("\u2022\t"), src_bullets))
        ok &= check("numbered lists converted",
                    bool(re.search(r"^\d+\.\t", out, re.M)) or src_numbered == 0)
    else:
        print("  SKIP  list checks - sample has no top-level markdown lists")

    # -- idempotency ------------------------------------------------------
    twice = P.run(body_lines(doc))
    stable = content_words(body_lines(twice)) == content_words(body_lines(doc))
    ok &= check("second run changes nothing", stable)

    print("RESULT: %s" % ("ALL PASS" if ok else "FAILURES PRESENT"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())