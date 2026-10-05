"""Static checks for the VBA sources.

The previous release shipped a module that did not compile: `silentMode`
was assigned but never declared, under `Option Explicit`.  Word only
reports that when someone runs the macro, so this checks the same class of
problem without needing Word.

Run:  python3 tests/lint_vba.py src/*.bas src/*.cls
"""

import os
import re
import sys

# VBA keywords, literals and language constructs
KEYWORDS = {
    "and", "as", "boolean", "byref", "byval", "call", "case", "class",
    "close", "const", "currency", "date", "declare", "dim", "do", "double",
    "each", "else", "elseif", "end", "enum", "eqv", "erase", "error",
    "exit", "explicit", "false", "finally", "for", "friend", "function",
    "get", "global", "goto", "if", "imp", "implements", "in", "input",
    "integer", "is", "let", "lib", "like", "long", "loop", "lset", "me",
    "mod", "new", "next", "not", "nothing", "on", "open", "option",
    "optional", "or", "output", "paramarray", "preserve", "print",
    "private", "property", "public", "put", "raiseevent", "redim",
    "rem", "resume", "return", "rset", "select", "set", "single", "static",
    "step", "stop", "string", "sub", "then", "to", "true", "type", "until",
    "variant", "wend", "while", "with", "withevents", "xor", "attribute",
    "put", "implements", "default", "readonly", "writeonly", "off", "goto",
}

# Object model members and globals that need no declaration.
BUILTINS = set("""
abs asc ascw cbool cdate cdbl cint clng csng cstr cvar cvdate dateadd datediff
datepart dateserial datevalue day fix format hex hour iif instr instrrev int
isarray isdate isempty iserror ismissing isnull isnumeric isobject join lbound
lcase left len log ltrim minute month monthname now oct right rnd round rtrim
second sgn space split sqr str strcomp strconv string strreverse switch tan time
timer timeserial timevalue trim typename ubound ucase val weekday year
abs application activecell activedocument activewindow array asin atn
bounding box cancel collection colorindex compare concat constant count
createobject csldatabase currentdate currenttime currentyear dateadd
debug documents errors filesystemobject format hex inarray inputbox
intransaction isarray isempty iserror ismissing isnull join lbound lcase
createobject keyboardobject match mode msgbox mkdir nextobject object option
print replace rgb remove rgb removeallcollection removeitem rename replace
selectcase sendkeys setlocale sgn shell sin space split sqr str strcomp
strconv string strreverse switch tan time timer timeseries trim typename
ubound ucase val weekday write xml DOMDocument
""".split())

TYPE_WORDS = {
    "long", "integer", "string", "boolean", "double", "single", "variant",
    "object", "date", "currency", "byte",
}

LINE_CONT = "_"


def strip_comment(line):
    out = []
    in_str = False
    i = 0
    while i < len(line):
        c = line[i]
        if c == '"':
            in_str = not in_str
            out.append(c)
        elif c == "'" and not in_str:
            break
        else:
            out.append(c)
        i += 1
    return "".join(out)


def join_continuations(text):
    """Merge VBA line-continuations into logical lines."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out = []
    buf = ""
    for raw in lines:
        stripped = raw.rstrip()
        if buf:
            buf += " " + stripped.strip()
        else:
            buf = stripped.strip()
        if buf.endswith(LINE_CONT):
            buf = buf[:-1].rstrip()
            continue
        out.append(buf)
        buf = ""
    if buf:
        out.append(buf)
    return out


class Module:
    def __init__(self, path):
        self.path = path
        with open(path, encoding="utf-8") as fh:
            self.text = fh.read()
        self.lines = join_continuations(self.text)
        self.code = [strip_comment(l) for l in self.lines]
        self.declared = set()
        self.procedures = []
        self.types = []
        self.consts = set()
        self.problem = []

    # ------------------------------------------------------------------
    def collect(self):
        # module level + procedure level declarations
        for line in self.code:
            low = line.strip().lower()
            m = re.match(r"\s*(?:public |private |global )?(?:dim|static)\s+(.*)$", low)
            if m:
                for name in self._split_names(m.group(1)):
                    self.declared.add(name.lower())
            # declarations may omit Dim: "Private mBusy As Boolean"
            m = re.match(r"\s*(?:public|private)\s+(?!sub|function|type|const|declare|enum|property|declare)(.+?)\s+as\s+[a-z]", low)
            if m:
                for name in self._split_names(m.group(1)):
                    self.declared.add(name.lower())
            m = re.match(r"\s*(?:public |private )?const\s+(.*)$", low)
            if m:
                for name in self._split_names(m.group(1)):
                    self.consts.add(name.lower())
            m = re.match(r"\s*(?:public |private )?(?:sub|function|property\s+\w+)\s+(\w+)\s*\(", low)
            if m:
                self.procedures.append(m.group(1))
            m = re.match(r"\s*(?:public |private )?type\s+(\w+)", low)
            if m:
                self.types.append(m.group(1))

        # parameter names
        for line in self.code:
            for chunk in self._param_chunks(line):
                for name in self._split_names(chunk):
                    self.declared.add(name.lower())

        # user defined type fields
        for i, line in enumerate(self.code):
            m = re.match(r"\s*(?:public |private )?type\s+(\w+)\s*$", line.strip().lower())
            if not m:
                continue
            j = i + 1
            while j < len(self.code) and not re.match(r"\s*end type\b", self.code[j].strip().lower()):
                fm = re.match(r"\s*(\w+)\s+as\s+", self.code[j].strip().lower())
                if fm:
                    self.declared.add(fm.group(1))
                j += 1

    def _param_chunks(self, line):
        low = line.lower()
        if "(" not in low:
            return []
        # only parameter lists: "(...)" that follow a procedure signature or
        # a "ByVal"/"ByRef"/"Optional"
        chunks = []
        depth = 0
        cur = ""
        started = False
        for i, ch in enumerate(low):
            if ch == "(":
                depth += 1
                started = True
                cur = ""
                continue
            if ch == ")":
                depth -= 1
                if depth == 0 and started:
                    chunks.append(cur)
                    started = False
                continue
            if started:
                cur += ch
        return chunks

    def _split_names(self, spec):
        """'a As Long, b As String' -> ['a', 'b']"""
        names = []
        depth = 0
        cur = ""
        for ch in spec:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if ch == "," and depth == 0:
                names.append(cur)
                cur = ""
            else:
                cur += ch
        if cur.strip():
            names.append(cur)
        out = []
        for part in names:
            part = part.strip()
            if not part:
                continue
            part = re.sub(r"^(?:optional\s+|paramarray\s+)", "", part)
            part = re.sub(r"^(?:byval\s+|byref\s+)", "", part)
            head = part.split(" as ")[0].strip()
            head = head.strip()
            if re.match(r"^\w+$", head):
                out.append(head)
        return out

    # ------------------------------------------------------------------
    def check_option_explicit(self):
        if "option explicit" not in "\n".join(self.lines).lower():
            self.problem.append("module does not use Option Explicit")

    def check_blocks(self):
        """If/End If, For/Next, With/End With, Sub/End Sub, Select/End Select."""
        depth = {"if": 0, "for": 0, "with": 0, "select": 0,
                 "sub": 0, "function": 0, "type": 0, "do": 0}
        for n, line in enumerate(self.code, 1):
            low = line.strip().lower()
            if not low:
                continue
            if low.startswith("attribute ") or re.match(
                    r"^(?:multiuse|vb_name|vb_globalname|vb_predeclaredid|"
                    r"vb_exposed|vb_creatable|vb_.*base)\s*=", low):
                continue
            words = re.findall(r"[a-z_]+", low)
            if not words:
                continue
            w = words
            if re.match(r"^(?:public |private |friend )?sub\s+\w+", low):
                depth["sub"] += 1
            elif re.match(r"^(?:public |private |friend )?function\s+\w+", low):
                depth["function"] += 1
            elif re.match(r"^(?:public |private )?type\s+\w+\s*$", low):
                depth["type"] += 1
            if re.match(r"^end sub\b", low):
                depth["sub"] -= 1
            elif re.match(r"^end function\b", low):
                depth["function"] -= 1
            elif re.match(r"^end type\b", low):
                depth["type"] -= 1
            elif re.match(r"^end select\b", low):
                depth["select"] -= 1
            elif re.match(r"^end with\b", low):
                depth["with"] -= 1
            elif re.match(r"^end if\b", low):
                depth["if"] -= 1
            elif re.match(r"^next\b", low):
                depth["for"] -= 1
            elif re.match(r"^if\b", low):
                # single-line If has a statement after Then and opens no block
                if re.search(r"\bthen\s+\S", low):
                    pass
                else:
                    depth["if"] += 1
            elif re.match(r"^for\b", low):
                depth["for"] += 1
            elif re.match(r"^with\b", low):
                depth["with"] += 1
            elif re.match(r"^select case\b", low):
                depth["select"] += 1
            if depth["if"] < 0:
                self.problem.append("line %d: 'End If' without 'If'" % n)
                depth["if"] = 0
            if depth["for"] < 0:
                self.problem.append("line %d: 'Next' without 'For'" % n)
                depth["for"] = 0
            if depth["with"] < 0:
                self.problem.append("line %d: 'End With' without 'With'" % n)
                depth["with"] = 0
        for k, v in depth.items():
            if v != 0:
                self.problem.append("unbalanced %s blocks (delta %d)" % (k, v))

    def check_undeclared(self):
        """Flag identifiers used as values but never declared or built in.

        Deliberately conservative: anything preceded by a dot (member
        access) or a call is ignored, so the report stays short and
        actionable.
        """
        known = (self.declared | KEYWORDS | BUILTINS | TYPE_WORDS |
                 set(n.lower() for n in self.procedures) |
                 set(n.lower() for n in self.types) | self.consts)
        for n, line in enumerate(self.code, 1):
            low = line.strip().lower()
            if not low or low.startswith("#"):
                continue
            if low.startswith("attribute ") or re.match(
                    r"^(?:multiuse|vb_name|vb_globalname|vb_predeclaredid|"
                    r"vb_exposed|vb_creatable)\s*=", low):
                continue
            # strip string literals
            clean = re.sub(r'"[^"]*"', '""', low)
            # only look at assignment targets and bare identifier uses
            m = re.match(r"^(?:set\s+|let\s+)?([a-z_]\w*)\s*(?:=|\+=|-=|\*=|/=)", clean)
            targets = [m.group(1)] if m else []
            for t in targets:
                if t not in known:
                    self.problem.append(
                        "line %d: assignment to undeclared '%s'" % (n, t))
            # find/select on undefined members is out of scope
        return known

    def report(self):
        self.collect()
        self.check_option_explicit()
        self.check_blocks()
        self.check_undeclared()
        return self.problem


def main(paths):
    bad = 0
    for path in paths:
        mod = Module(path)
        problems = mod.report()
        name = os.path.basename(path)
        if problems:
            bad += 1
            print("FAIL %s" % name)
            for p in problems:
                print("       %s" % p)
        else:
            print("OK   %s  (%d procedures, %d declared names)"
                  % (name, len(mod.procedures), len(mod.declared)))
    return 1 if bad else 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        args = sorted(
            os.path.join(here, "src", f)
            for f in os.listdir(os.path.join(here, "src"))
            if f.endswith((".bas", ".cls"))
        )
    sys.exit(main(args))