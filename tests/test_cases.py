"""Edge-case tests for the pipeline model.

The real sample is a specification: headings plus fenced TSV tables.  These
cases cover the constructs it does not contain, plus the interactions that
were broken in the old VBA.
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pipeline_ref as P

CASES = []


def case(fn):
    CASES.append(fn)
    return fn


def text_of(doc):
    """Flattened output, table cells joined with pipes."""
    out = []
    for line in doc.lines:
        m = re.match(r"^\x00T(\d+)\x00$", line)
        if m:
            out.extend("|".join(r) for r in doc.tables[int(m.group(1))])
        else:
            out.append(line)
    return out


def fmt(lines):
    return text_of(P.run(lines))


def eq(got, want, label):
    if got == want:
        print("  PASS  " + label)
        return True
    print("  FAIL  " + label)
    print("        got  %r" % (got,))
    print("        want %r" % (want,))
    return False


@case
def t_atx_headings():
    return eq(fmt(["# One", "## Two", "### Three", "body"]),
              ["One", "Two", "Three", "body"], "ATX heading markers removed")


@case
def t_setext_underlined():
    out = fmt(["Title", "=====", "Sub", "-----", "body"])
    return eq(out, ["Title", "Sub", "body"], "setext underlines become headings")


@case
def t_hr_not_setext():
    out = fmt(["para", "", "---", "", "next"])
    return eq(out, ["para", "", "", "next"], "standalone --- is a rule, not a heading")


@case
def t_setext_vs_bullet():
    out = fmt(["- item one", "---"])
    return eq(out, ["\u2022\titem one"], "--- after a bullet is a rule, not setext")


@case
def t_bullets():
    return eq(fmt(["- a", "* b", "+ c"]),
              ["\u2022\ta", "\u2022\tb", "\u2022\tc"], "bullet glyphs")


@case
def t_numbered():
    return eq(fmt(["1. a", "2. b", "10. c"]),
              ["1.\ta", "2.\tb", "10.\tc"], "numbered lists keep their numbers")


@case
def t_blockquote():
    # The quote is carried by the paragraph role, not by a literal glyph, so
    # the text itself only loses the "> " marker.
    roles = P.process_plain_run(["> quoted text"])
    ok = roles == [("q", "quoted text")]
    return eq(ok, True, "blockquote marker stripped and role recorded")


@case
def t_code_block_kept():
    out = fmt(["before", "```python", "x = 1", "```", "after"])
    return eq(out, ["before", "x = 1", "after"], "code fence unwrapped, body kept")


@case
def t_code_block_untouched_by_inline():
    out = fmt(["```", "keep **this** raw", "```"])
    return eq(out, ["keep **this** raw"], "markdown inside code is left alone")


@case
def t_unterminated_fence():
    out = fmt(["intro", "```", "still body"])
    return eq(out, ["intro", "still body"], "dangling fence loses only its marker")


@case
def t_tsv_fence_to_table():
    doc = P.run(["```tsv", "a\tb", "1\t2", "```"])
    ok = len(doc.tables) == 1 and doc.tables[0] == [["a", "b"], ["1", "2"]]
    return eq(ok, True, "fenced tsv becomes a real table")


@case
def t_tsv_fence_with_prose():
    doc = P.run(["```tsv", "note about the table", "a\tb", "1\t2", "```"])
    ok = len(doc.tables) == 1 and doc.tables[0] == [["a", "b"], ["1", "2"]]
    ok &= "note about the table" in doc.lines
    return eq(ok, True, "prose inside a tsv fence is not swallowed by the table")


@case
def t_tsv_ragged_width():
    doc = P.run(["```tsv", "a\tb\tc", "1\t2", "```"])
    ok = len(doc.tables) == 1 and doc.tables[0][1] == ["1", "2", ""]
    return eq(ok, True, "short rows are padded, not dropped")


@case
def t_plain_tabs_outside_fence():
    doc = P.run(["a\tb", "1\t2"])
    return eq(len(doc.tables), 1, "unfenced tab rows still become a table")


@case
def t_inline_styles():
    return eq(fmt(["plain **bold** and *em* and ~~gone~~"]),
              ["plain bold and em and gone"], "inline markers stripped")


@case
def t_inline_code():
    return eq(fmt(["call `foo(1)` now"]), ["call foo(1) now"], "inline code unwrapped")


@case
def t_link():
    return eq(fmt(["see [docs](http://x.y) here"]),
              ["see docs here"], "markdown links unwrapped")


@case
def t_idempotent():
    src = ["# H", "- a", "> q", "1. n", "```tsv", "a\tb", "```", "text **b**"]
    once = fmt(src)
    return eq(fmt(once), once, "running twice is a no-op")


@case
def t_empty_and_blank():
    return eq(fmt([]), [], "empty document")


@case
def t_thematic_break_spaced():
    # Markdown treats "- - -" as a thematic break, not a list item.
    out = fmt(["- - -"])
    return eq(out, [], "spaced dashes are a thematic break")


@case
def t_headings_around_table():
    src = ["# Title", "```tsv", "a\tb", "1\t2", "```", "tail"]
    out = fmt(src)
    return eq(out, ["Title", "a|b", "1|2", "tail"], "table sits between headings")


@case
def t_crlf_like_trailing_space():
    return eq(fmt(["# H   ", "   "]), ["H", "   "], "trailing spaces trimmed")


def main():
    failed = 0
    for fn in CASES:
        print(fn.__name__)
        if not fn():
            failed += 1
        print()
    print("%d/%d cases passed" % (len(CASES) - failed, len(CASES)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())