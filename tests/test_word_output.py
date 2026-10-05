"""Compare a docx that the Word add-in already formatted against the model.

Usage:
    python3 tests/test_word_output.py <original-markdown.docx> <formatted.docx>

The macro must have been run on a copy of the original, then saved.  Every
formatting bug found so far (stale offsets after the fence delete, tables
being re-scanned on a second run, non-short-circuit And) showed up here as
a shape or content mismatch that the pure model tests could not see.
"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pipeline_ref as P

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def ptext(el):
    out = []
    for n in el.iter():
        if n.tag == W + "t":
            out.append(n.text or "")
        elif n.tag == W + "tab":
            out.append("\t")
    return "".join(out)


def body_lines_of(path):
    """Body children in document order, with tables expanded like the model."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    lines, tables = [], []
    for child in body:
        if child.tag == W + "p":
            lines.append(ptext(child))
        elif child.tag == W + "tbl":
            grid = child.find(W + "tblGrid")
            ncol = len(grid.findall(W + "gridCol")) if grid is not None else 0
            rows = []
            for tr in child.findall(W + "tr"):
                cells = [ptext(tc).strip() for tc in tr.findall(W + "tc")]
                cells += [""] * (ncol - len(cells))
                rows.append(cells[:ncol])
            tables.append(rows)
            lines.append("\x00T%d\x00" % (len(tables) - 1))
    return lines, tables


def all_lines(lines, tables):
    out = []
    for l in lines:
        m = re.match(r"^\x00T(\d+)\x00$", l)
        if m:
            out.extend("|".join(r) for r in tables[int(m.group(1))])
        else:
            out.append(l)
    return out


def words(lines):
    j = "\n".join(lines)
    j = re.sub(r"[`*_~>#|=\-]", "", j)
    return re.sub(r"\s+", "", j)


def main():
    md_path, word_path = sys.argv[1], sys.argv[2]
    md_lines, _ = body_lines_of(md_path)
    doc = P.run(md_lines)
    w_lines, w_tables = body_lines_of(word_path)

    ok = True

    def check(name, cond, detail=""):
        nonlocal ok
        ok &= bool(cond)
        print(("  PASS  " if cond else "  FAIL  ") + name + ("  -- " + detail if detail else ""))

    print("model : %d lines, %d tables" % (len(doc.lines), len(doc.tables)))
    print("word  : %d lines, %d tables" % (len(w_lines), len(w_tables)))

    def in_doc_order(lines, tables):
        order = [int(m.group(1)) for m in
                 (re.match(r"^\x00T(\d+)\x00$", l) for l in lines) if m]
        return [tables[i] for i in order]

    m_tables = in_doc_order(doc.lines, doc.tables)

    check("table count matches", len(m_tables) == len(w_tables),
          "model=%d word=%d" % (len(m_tables), len(w_tables)))

    for i, (m, w) in enumerate(zip(m_tables, w_tables)):
        if len(m) != len(w) or any(len(a) != len(b) for a, b in zip(m, w)):
            check("table %d shape matches" % (i + 1), False,
                  "model %dx%s word %dx%s" % (len(m), len(m[0]) if m else 0, len(w), len(w[0]) if w else 0))
            break
    else:
        check("every table shape matches model", True)

    mw, ww = words(all_lines(doc.lines, doc.tables)), words(all_lines(w_lines, w_tables))
    check("no text lost or invented", mw == ww,
          "model=%d word=%d" % (len(mw), len(ww)))
    if mw != ww:
        for i, (a, b) in enumerate(zip(mw, ww)):
            if a != b:
                print("  first divergence at char %d" % i)
                print("    model: ...%s..." % mw[max(0, i - 40):i + 40])
                print("    word : ...%s..." % ww[max(0, i - 40):i + 40])
                break

    check("no fence markers survive", not [l for l in w_lines if "```" in l])
    check("no ATX headings survive", not [l for l in w_lines if P.ATX.match(l.strip())])
    check("no setext underlines survive",
          not re.search(r"^={2,}$", "\n".join(w_lines), re.M))

    w_body = [l for l in w_lines if not re.match(r"^\x00T\d+\x00$", l)]
    twice = P.run(w_body)
    check("model leaves the Word body alone", len(twice.tables) == 0,
          "model wanted %d tables" % len(twice.tables))
    check("model leaves the Word body text alone",
          words(all_lines(twice.lines, twice.tables)) == words(w_body))

    print("RESULT: %s" % ("MATCH" if ok else "MISMATCH"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())