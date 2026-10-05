"""Reference model of the ChatFormatter pipeline.

An executable description of the algorithm implemented by
src/ChatFormatter.bas, so the pipeline can be tested against the real
sample document without running VBA.

The design has three parts:

  1. scan()   - ONE forward pass that classifies every paragraph and
                groups consecutive paragraphs into typed blocks.  This is
                the only pass that touches the paragraph list, which is
                what makes the formatter linear instead of quadratic
                (the old code indexed doc.Paragraphs(i) inside loops).
  2. run()    - walks the block list backwards and rewrites each block.
                Going backwards means an edit never invalidates the
                position of a block that has not been visited yet.
  3. counters - every paragraph touch is counted, so a test can assert the
                total work stays proportional to the input size.

A block is (kind, start, end) where end is exclusive:
    plain   - ordinary markdown text
    code    - a fenced block with a non-TSV language
    tsv     - a fenced block tagged tsv/tab/table whose body is real TSV
    table   - a markdown pipe table
"""

import re

CODE_FENCE = "```"
TSV_LANGS = {"tsv", "tab", "tabs", "table"}
ATX = re.compile(r"^(#{1,5}) (.+)$")
NUMBERED = re.compile(r"^(\d{1,3})\. (.+)$")
BULLET = re.compile(r"^([-*+]) (.+)$")
QUOTE_LINE = re.compile(r"^> ?(.*)$")
SETEXT_EQ = re.compile(r"^={2,}$")
SETEXT_DASH = re.compile(r"^-{2,}$")

# Inline stage patterns.  Order matters: links and code first so their
# payloads are not re-read as emphasis, bold before italic so the asterisks
# left behind by ** are not treated as a second emphasis run.
BOLD = re.compile(r"\*\*(.+?)\*\*")
STRIKE = re.compile(r"~~(.+?)~~")
INLINE_CODE = re.compile(r"`([^`]+)`")
LINK = re.compile(r"\[([^\]]+)\]\([^)]*\)")
ITALIC = re.compile(r"(?<!\*)\*([^*]+)\*(?!\*)")


def inline(text):
    text = LINK.sub(r"\1", text)
    text = INLINE_CODE.sub(r"\1", text)
    text = BOLD.sub(r"\1", text)
    text = STRIKE.sub(r"\1", text)
    text = ITALIC.sub(r"\1", text)
    return text
BLOCKQUOTE = re.compile(r"^> ?(.*)$")
SETEXT_EQ = re.compile(r"^={2,}$")
SETEXT_DASH = re.compile(r"^-{2,}$")


class Block:
    __slots__ = ("kind", "start", "end", "lang", "closed")

    def __init__(self, kind, start, end, lang="", closed=True):
        self.kind = kind
        self.start = start
        self.end = end
        self.lang = lang
        self.closed = closed

    def __repr__(self):
        return "Block(%s,%d,%d,%r)" % (self.kind, self.start, self.end, self.lang)


def fence_info(line):
    """Return the language of a fence line, or None if it is not a fence."""
    s = line.strip()
    if not s.startswith(CODE_FENCE):
        return None
    rest = s[3:].strip()
    if "`" in rest:
        return None
    return rest.lower()


def scan(lines):
    """One forward pass.  Returns a list of Blocks covering every line."""
    blocks = []
    i = 0
    n = len(lines)
    plain_start = 0

    def flush(end):
        if end > plain_start:
            blocks.append(Block("plain", plain_start, end))

    while i < n:
        line = lines[i].strip()
        lang = fence_info(line)
        if lang is not None:
            close = -1
            j = i + 1
            while j < n:
                if fence_info(lines[j]) is not None:
                    close = j
                    break
                j += 1
            if close == -1:
                # Unterminated fence: keep the body as a code block but drop
                # the dangling marker.
                flush(i)
                blocks.append(Block("code", i, n, lang, closed=False))
                return blocks
            flush(i)
            body = lines[i + 1:close]
            kind = "tsv" if lang in TSV_LANGS and looks_like_tsv(body) else "code"
            blocks.append(Block(kind, i, close + 1, lang))
            i = close + 1
            plain_start = i
            continue
        i += 1
    flush(n)
    return blocks


def looks_like_tsv(rows):
    """A fenced `tsv` block counts as tabular if any body line has a tab.

    Real AI output is not perfectly uniform: a fenced block often mixes
    tab-separated rows with an occasional prose line.  Requiring every row
    to agree (the earlier rule) demoted whole blocks to plain code blocks,
    which is exactly the "my table did not convert" bug.
    """
    return any(r.count("\t") >= 1 and not LIST_MARKED.match(r) for r in rows)


# A line this formatter already produced carries a list marker followed by a
# tab.  It must not be re-read as a TSV data row on the next pass, otherwise
# every bullet turns into a one-cell table on the second run.
LIST_MARKED = re.compile(r"^\s*(?:[\u2022\-\*\+]\t|\d{1,3}\.\t)")


def tabular_runs(rows):
    """Split a fenced tsv body into (is_table, lines) runs.

    Consecutive lines that contain at least one tab become one table run;
    lines without tabs stay as ordinary paragraphs.
    """
    runs = []
    cur, cur_tab = [], None
    for r in rows:
        has_tab = (r.count("\t") >= 1 and bool(r.strip())
                   and not LIST_MARKED.match(r))
        if cur_tab is None or has_tab == cur_tab:
            cur.append(r)
            cur_tab = has_tab
        else:
            runs.append((cur_tab, cur))
            cur, cur_tab = [r], has_tab
    if cur:
        runs.append((cur_tab, cur))
    return runs


def build_grid(rows):
    rows = [r for r in rows if r.strip() and not LIST_MARKED.match(r)]
    cols = max(r.count("\t") for r in rows) + 1
    grid = []
    for row in rows:
        cells = [c.strip() for c in row.split("\t")]
        cells += [""] * (cols - len(cells))
        grid.append(cells[:cols])
    return grid


def is_pipe_row(s):
    s = s.strip()
    return len(s) >= 3 and s.startswith("|") and s.endswith("|") and s.count("|") >= 2


def is_separator(s):
    s = s.strip()
    if not s or "|" not in s:
        return False
    if "-" not in s:
        return False
    return all(c in "|-: \t" for c in s)


def pipe_grid(rows):
    body = [r.strip() for r in rows[2:]]
    cells = []
    for row in [rows[0]] + body:
        s = row.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        cells.append([c.strip() for c in s.split("|")])
    cols = max(len(r) for r in cells)
    for r in cells:
        r += [""] * (cols - len(r))
    return cells


def is_all_of(s, ch):
    return len(s) >= 2 and all(c == ch for c in s)


class Doc:
    def __init__(self, lines):
        self.lines = list(lines)
        self.ops = 0
        self.tables = []

    def touch(self):
        self.ops += 1

    def replace_block(self, start, end, new_lines):
        self.ops += 1
        self.lines[start:end] = new_lines

    def insert_table(self, at, grid):
        self.ops += 1
        self.tables.append(grid)
        self.lines[at:at] = ["\x00T%d\x00" % (len(self.tables) - 1)]


def process_plain_run(chunk):
    """Block-level + inline formatting for a run with no tabs and no fences.

    Setext underlines are resolved before the generic horizontal-rule test:
    "Title" followed by "-----" is a heading, while a "---" that follows a
    list item or sits on its own is a rule.
    """
    out = []
    for raw in chunk:
        t = raw.strip()
        if not t:
            out.append(("p", raw))
            continue

        if out and SETEXT_EQ.match(t):
            out[-1] = ("h1", out[-1][1])
            continue
        if out and SETEXT_DASH.match(t) and not BULLET.match(out[-1][1].strip()):
            out[-1] = ("h2", out[-1][1])
            continue

        m = ATX.match(t)
        if m:
            out.append(("h%d" % len(m.group(1)), inline(m.group(2))))
            continue
        if len(t) >= 3 and is_hr(t):
            continue
        bq = QUOTE_LINE.match(t)
        if bq:
            out.append(("q", inline(bq.group(1))))
            continue
        bu = BULLET.match(t)
        if bu:
            out.append(("ul", "\u2022\t" + inline(bu.group(2))))
            continue
        nu = NUMBERED.match(t)
        if nu:
            out.append(("ol", nu.group(1) + ".\t" + inline(nu.group(2))))
            continue
        out.append(("p", inline(t)))

    return out


def run(lines):
    doc = Doc(lines)
    blocks = scan(doc.lines)
    doc.touch()  # the scan itself is linear

    # Backwards: an edit never shifts the position of an unvisited block.
    for b in reversed(blocks):
        doc.touch()
        chunk = doc.lines[b.start:b.end]
        if b.kind == "tsv":
            out = []
            for is_tab, run in tabular_runs(chunk[1:-1]):
                if is_tab:
                    grid = build_grid(run)
                    out.append("\x00T%d\x00" % len(doc.tables))
                    doc.tables.append(grid)
                else:
                    out.extend(run)
            doc.replace_block(b.start, b.end, out)
        elif b.kind == "code":
            # Drop the opening fence, and the closing fence only when there
            # actually was one - an unterminated fence keeps its last line.
            doc.replace_block(b.start, b.end,
                              list(chunk[1:-1]) if b.closed else list(chunk[1:]))
        else:
            # Tab runs become tables *before* list markers are inserted,
            # otherwise the tab in "1.\titem" looks like a table row.
            out = []
            for is_tab, run in tabular_runs(chunk):
                if is_tab:
                    grid = build_grid(run)
                    out.append("\x00T%d\x00" % len(doc.tables))
                    doc.tables.append(grid)
                else:
                    out.extend(t for _, t in process_plain_run(run))
            doc.replace_block(b.start, b.end, out)

    return doc

def is_bullet_line(s):
    return bool(re.match(r"^([-*+]) ", s))


def is_hr(t):
    core = t
    for ch in "-*_=~":
        core = core.replace(ch, "")
    core = core.replace("`", "").strip()
    return len(t) >= 3 and core == ""