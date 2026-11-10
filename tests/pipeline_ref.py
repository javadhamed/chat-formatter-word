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
DIGIT_RE = r"[0-9\u06f0-\u06f9\u0660-\u0669]"
ATX = re.compile(r"^(#{1,5}) (.+)$")
NUMBERED = re.compile(rf"^({DIGIT_RE}{{1,3}})\. (.+)$")
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

UNICODE_DASHES = "\u2012\u2013\u2014\u2015\u2500"
SENTENCE_PUNCT = set(".?!:;)]\"'\u061f\u061b\u060c\u00bb\u2014\u2013")


def digit_value(ch):
    code = ord(ch)
    if 48 <= code <= 57:
        return code - 48
    if 0x06F0 <= code <= 0x06F9:
        return code - 0x06F0
    if 0x0660 <= code <= 0x0669:
        return code - 0x0660
    return -1


def parse_digits(s):
    acc = 0
    for ch in s:
        d = digit_value(ch)
        if d < 0:
            return -1
        acc = acc * 10 + d
    return acc


def normalize_dashes(s):
    if not s:
        return s
    out = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "`":
            close = s.find("`", i + 1)
            if close != -1:
                out.append(s[i:close + 1])
                i = close + 1
                continue
        elif s[i:i + 2] == "](":
            close = s.find(")", i + 2)
            if close != -1:
                out.append(s[i:close + 1])
                i = close + 1
                continue

        if ch in UNICODE_DASHES or ch == "-":
            has_uni = ch in UNICODE_DASHES
            j = i + 1
            while j < n and (s[j] in UNICODE_DASHES or s[j] == "-"):
                if s[j] in UNICODE_DASHES:
                    has_uni = True
                j += 1
            if has_uni:
                out.append("-")
            elif j - i >= 2:
                prev_ch = s[i - 1] if i > 0 else ""
                next_ch = s[j] if j < n else ""
                if prev_ch == " " and next_ch == " ":
                    out.append("-")
                elif prev_ch and next_ch and prev_ch not in " \t" and next_ch not in " \t":
                    out.append("-")
                else:
                    out.append(s[i:j])
            else:
                out.append(ch)
            i = j
            continue

        out.append(ch)
        i += 1
    return "".join(out)


def split_sequential_numbered(s):
    n = len(s)
    if n < 6 or ". " not in s:
        return [s]

    cands = []
    i = 0
    while i < n:
        if s[i] == "`":
            close = s.find("`", i + 1)
            if close != -1:
                i = close + 1
                continue
        elif s[i:i + 2] == "](":
            close = s.find(")", i + 2)
            if close != -1:
                i = close + 1
                continue

        prev_ch = s[i - 1] if i > 0 else ""
        if i == 0 or prev_ch in " \t":
            dlen = 0
            while i + dlen < n and dlen < 4 and digit_value(s[i + dlen]) >= 0:
                dlen += 1
            if 1 <= dlen <= 3 and i + dlen + 1 < n and s[i + dlen:i + dlen + 2] == ". ":
                if s[i + dlen + 2:].strip():
                    num_val = parse_digits(s[i:i + dlen])
                    is_punct = False
                    is_colon = False
                    if i > 0:
                        j = i - 1
                        while j >= 0 and s[j] in " \t":
                            j -= 1
                        if j >= 0:
                            is_colon = (s[j] == ":")
                            is_punct = (s[j] in SENTENCE_PUNCT)
                    cands.append((i, num_val, is_punct, is_colon))
                    i += dlen + 2
                    continue
        i += 1

    if len(cands) < 2:
        return [s]

    start_idx = -1
    if cands[0][0] == 0 and 1 <= cands[0][1] <= 50:
        start_idx = 0
    else:
        for idx, (_, num_val, _, is_colon) in enumerate(cands):
            if num_val == 1 and is_colon:
                start_idx = idx
                break
    if start_idx == -1:
        return [s]

    chain = [start_idx]
    want = cands[start_idx][1] + 1
    all_punct = True
    for idx in range(start_idx + 1, len(cands)):
        if cands[idx][1] == want:
            chain.append(idx)
            if not cands[idx][2]:
                all_punct = False
            want += 1

    if len(chain) < 2:
        return [s]
    if len(chain) == 2 and (cands[chain[0]][1] != 1 or not all_punct):
        return [s]

    out = []
    first_pos = cands[chain[0]][0]
    if first_pos > 0:
        intro = s[:first_pos].strip()
        if intro:
            out.append(intro)
    for k, cidx in enumerate(chain):
        seg_from = cands[cidx][0]
        seg_to = cands[chain[k + 1]][0] if k + 1 < len(chain) else n
        piece = s[seg_from:seg_to].strip()
        if piece:
            out.append(piece)
    return out if len(out) > 1 else [s]


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
        nonlocal plain_start
        k = plain_start
        seg_start = k
        while k < end:
            pipe_end = find_pipe_table_end(lines, k, end)
            if pipe_end > k:
                if k > seg_start:
                    blocks.append(Block("plain", seg_start, k))
                blocks.append(Block("table", k, pipe_end))
                k = pipe_end
                seg_start = k
            else:
                k += 1
        if end > seg_start:
            blocks.append(Block("plain", seg_start, end))

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
LIST_MARKED = re.compile(rf"^\s*(?:[\u2022\-\*\+]\t|{DIGIT_RE}{{1,3}}\.\t)")


def tabular_runs(rows, min_rows=1):
    """Split a body into (is_table, lines) runs.

    Consecutive lines that contain at least one tab become one table run
    when the run has at least min_rows non-empty lines with >= 2 columns.
    """
    raw_runs = []
    cur, cur_tab = [], None
    for r in rows:
        has_tab = (r.count("\t") >= 1 and bool(r.strip())
                   and not LIST_MARKED.match(r))
        if cur_tab is None or has_tab == cur_tab:
            cur.append(r)
            cur_tab = has_tab
        else:
            raw_runs.append((cur_tab, cur))
            cur, cur_tab = [r], has_tab
    if cur:
        raw_runs.append((cur_tab, cur))

    runs = []
    for is_tab, run in raw_runs:
        if is_tab:
            non_empty = [x for x in run if x.strip() and not LIST_MARKED.match(x)]
            max_cols = max((x.count("\t") + 1 for x in non_empty), default=0)
            if len(non_empty) < min_rows or max_cols < 2:
                is_tab = False
        if runs and not is_tab and not runs[-1][0]:
            runs[-1][1].extend(run)
        else:
            runs.append((is_tab, list(run)))
    return runs


def build_grid(rows):
    rows = [r for r in rows if r.strip() and not LIST_MARKED.match(r)]
    cols = max(r.count("\t") for r in rows) + 1
    grid = []
    for row in rows:
        cells = [inline(normalize_dashes(c.strip())) for c in row.split("\t")]
        cells += [""] * (cols - len(cells))
        grid.append(cells[:cols])
    return grid


def pipe_col_count(s):
    t = s.strip()
    if len(t) < 2 or "|" not in t:
        return 0
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|"):
        t = t[:-1]
    if not t:
        return 0
    return t.count("|") + 1


def is_pipe_row(s):
    return "|" in s and pipe_col_count(s) >= 2


def is_bordered_pipe_row(s):
    t = s.strip()
    return len(t) >= 3 and t.startswith("|") and t.endswith("|") and t.count("|") >= 2


def is_separator(s):
    t = s.strip()
    if len(t) < 3 or "|" not in t or "-" not in t:
        return False
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|"):
        t = t[:-1]
    parts = t.split("|")
    if len(parts) < 2:
        return False
    for p in parts:
        c = p.strip()
        if not c or "-" not in c or not all(ch in "-:" for ch in c):
            return False
    return True


def find_pipe_table_end(lines, idx, last):
    s0 = lines[idx].strip()
    if not is_pipe_row(s0):
        return 0
    if idx + 1 < last and is_separator(lines[idx + 1]):
        j = idx + 2
        while j < last:
            sj = lines[j].strip()
            if not sj or fence_info(sj) is not None or not is_pipe_row(sj) or is_separator(sj):
                break
            j += 1
        return j
    if is_bordered_pipe_row(s0) and not is_separator(s0):
        j = idx + 1
        while j < last:
            sj = lines[j].strip()
            if not sj or fence_info(sj) is not None:
                break
            if is_bordered_pipe_row(sj) and not is_separator(sj):
                j += 1
            else:
                break
        if j > idx + 1:
            return j
    return 0


def pipe_grid(rows):
    data_rows = [r.strip() for r in rows if r.strip() and not is_separator(r)]
    cells = []
    for row in data_rows:
        s = row.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        cells.append([inline(normalize_dashes(c.strip())) for c in s.split("|")])
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
            out.append(("h%d" % len(m.group(1)), inline(normalize_dashes(m.group(2)))))
            continue
        if len(t) >= 3 and is_hr(t):
            continue
        bq = QUOTE_LINE.match(t)
        if bq:
            out.append(("q", inline(normalize_dashes(bq.group(1)))))
            continue
        bu = BULLET.match(t)
        if bu:
            out.append(("ul", "\u2022\t" + inline(normalize_dashes(bu.group(2)))))
            continue
        pieces = split_sequential_numbered(t)
        for piece in pieces:
            nu = NUMBERED.match(piece)
            if nu:
                out.append(("ol", nu.group(1) + ".\t" + inline(normalize_dashes(nu.group(2)))))
            else:
                out.append(("p", inline(normalize_dashes(piece))))

    return out


def expand_soft_breaks(lines):
    out = []
    for ln in lines:
        if "\x0b" in ln:
            out.extend(ln.split("\x0b"))
        else:
            out.append(ln)
    return out


def run(lines):
    doc = Doc(expand_soft_breaks(lines))
    blocks = scan(doc.lines)
    doc.touch()  # the scan itself is linear

    # Backwards: an edit never shifts the position of an unvisited block.
    for b in reversed(blocks):
        doc.touch()
        chunk = doc.lines[b.start:b.end]
        if b.kind == "tsv":
            out = []
            for is_tab, sub in tabular_runs(chunk[1:-1], min_rows=1):
                if is_tab:
                    grid = build_grid(sub)
                    out.append("\x00T%d\x00" % len(doc.tables))
                    doc.tables.append(grid)
                else:
                    out.extend(sub)
            doc.replace_block(b.start, b.end, out)
        elif b.kind == "table":
            grid = pipe_grid(chunk)
            doc.replace_block(b.start, b.end, ["\x00T%d\x00" % len(doc.tables)])
            doc.tables.append(grid)
        elif b.kind == "code":
            # Drop the opening fence, and the closing fence only when there
            # actually was one - an unterminated fence keeps its last line.
            doc.replace_block(b.start, b.end,
                              list(chunk[1:-1]) if b.closed else list(chunk[1:]))
        else:
            # Tab runs become tables *before* list markers are inserted,
            # otherwise the tab in "1.\titem" looks like a table row.
            out = []
            for is_tab, sub in tabular_runs(chunk, min_rows=2):
                if is_tab:
                    grid = build_grid(sub)
                    out.append("\x00T%d\x00" % len(doc.tables))
                    doc.tables.append(grid)
                else:
                    out.extend(t for _, t in process_plain_run(sub))
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