"""
Static Option Explicit check for a VBA standard module.

VBA compiles a procedure the first time it runs, so a build gate that only
calls macros can never see an undeclared identifier in a procedure that the
test never reaches. When the same module is loaded as a global template Word
compiles all of it and the build breaks with "Compile error in hidden module",
with no line number. This finds those names before Word ever sees the file.

Method: collect every name the module declares (module level, procedure level,
parameters, For/With/Do variables), then flag every bare identifier that is not
declared and is not a known VBA or Word built-in. Member accesses after a dot
are skipped, since those resolve against the object model.
"""
import re
import sys

BUILTINS = set("""
addressof alias and any app activate application array as asc atn attribute boolean byref byte
byval call case cbool cdate cdbl cint clng close compare const csng cstr currency date
debug decimal declare defbool defbyte defcur defdate defdbl defdecndefint deflng defobj
defsng defstr deftype defvar dim dir do double each else elseif empty end enum eqv erase
error event exit explicit false for friend function get global gosub goto if imp implements
in integer is let lib like long loop lset ltrim me mid mod new next not nothing null
on open option optional or preserve print private property public raiseevent randomize
redim rem resume return rset selectcase set single static step stop string sub then to
true type typeof until variant wend while with withevents xor vbcr vbcrlf vbformfeed
newline vbtab vby vbnullstring vbtrue vbfalse error resume goto
abs asc cbool cdate cdbl cint clng csng cstr fix int sgn sqr val
chr chrw dateadd datediff datepart dateserial datevalue day format hour isdate
minute month now second time timer timeserial timevalue weekday year
array lbound ubound isarray isnull isnumeric isempty isobject isdate ismissing
lcase ucase left right mid trim space strreverse string len instr instrrev
replace split join format strcomp strreverse
rnd randomize
iif switch partition
msgbox inputbox beep
environ curvbvar cvar
createobject getobject
doevents
dir kill filecopy mkdir rmdir name freeresume
load savefield loaddata
set let
redim preserve
erase
typeof vartype vartype
end enum
open close print write input line output append random binary
lock unlock seek
iif
preserve
typename vba
debug
inputbox
new
end
""".split())

# Names that come from the Word / Office object model and are legal even when
# they are not declared in the module.
WORD_GLOBALS = set("""
application activedocument activeselection documents windows selection
range table paragraph paragraphs rows columns cells cell row column
style styles font fonts listformat lists paragraphs options
wordbasic vbaworkspace
""" .split())

decl_re = re.compile(
    r'^\s*(?:(?:public|private|friend|dim|static|global)\s+)*'
    r'(?:with\s+events\s+)?'
    r'([A-Za-z_]\w*)\s*(?:\(\s*([^)]*)\s*\))?\s*(?:as\s+\S+)?\s*$',
    re.I)

comment_re = re.compile(r'^\s*(?:rem\b|)')
string_re = re.compile(r'"(?:[^"]|"")*"')

# A line that opens a declaration, e.g. "Dim a, b As Long, c"
def declared_names_on(line):
    line = re.sub(r'"(?:[^"]|"")*"', '""', line)
    if comment_re.match(line):
        return []
    out = []
    # strip any trailing comment
    line = re.split(r'\s\'\s', line)[0]
    m = re.match(r'^\s*(?:public|private|friend|global|dim|static)\b(?!\s*sub)(?!\s*function)(.*)$', line, re.I)
    if not m:
        return []
    body = m.group(1)
    # cut at "As <type>" for the tail item
    for part in re.split(r',', body):
        part = part.strip()
        if not part:
            continue
        nm = re.match(r'^([A-Za-z_]\w*)', part)
        if nm:
            out.append(nm.group(1))
    return out


def main(path):
    raw = open(path, encoding='utf-8', errors='replace').read()
    lines = raw.splitlines()

    declared = set()
    # module level and procedure level declarations
    for ln in lines:
        for nm in declared_names_on(ln):
            declared.add(nm.lower())

    # procedure signatures, including params
    for m in re.finditer(r'^\s*(?:public|private|friend)?\s*(?:sub|function)\s+(\w+)\s*\(([^)]*)\)', raw, re.I | re.M):
        declared.add(m.group(1).lower())
        for part in m.group(2).split(','):
            part = re.sub(r'\b(as\s+\w+)?\s*$', '', part.strip(), flags=re.I).strip()
            nm = re.match(r'([A-Za-z_]\w*)', part)
            if nm:
                declared.add(nm.group(1).lower())

    # For Each x, With x, Do While (x) etc.
    for m in re.finditer(r'\bfor\s+each\s+([A-Za-z_]\w*)|\bwith\s+([A-Za-z_]\w*)|\bset\s+([A-Za-z_]\w*)\s*=', raw, re.I):
        for g in m.groups():
            if g:
                declared.add(g.lower())

    # Enum members and Const/Type names
    for m in re.finditer(r'^\s*(?:public|private)?\s*const\s+([A-Za-z_]\w*)|^Public\s+(?:Sub|Function|Type|Enum)\s+(\w+)', raw, re.I | re.M):
        for g in m.groups():
            if g:
                declared.add(g.lower())

    known = declared | BUILTINS | WORD_GLOBALS

    problems = []
    in_block = False
    for no, ln in enumerate(lines, 1):
        code = string_re.sub('""', ln)
        code = re.sub(r"'.*$", '', code)
        if not code.strip():
            continue
        # skip declaration-ish lines: their identifiers are names, not uses
        if re.match(r'^\s*(public|private|friend|global|dim|static)\b', code, re.I) \
           and not re.search(r'\bthen\b|\bAs\b\s*$', code, re.I) is None:
            pass
        if re.match(r'^\s*(public|private|friend|global|dim|static)\b', code, re.I) \
           and not re.search(r'\bthen\b', code, re.I) \
           and not re.search(r'\busing\b', code, re.I):
            # a pure declaration line still contains no bare uses
            pass
        for m in re.finditer(r'(?<![.\w])([A-Za-z_]\w*)', code):
            nm = m.group(1)
            low = nm.lower()
            if low in known:
                continue
            if re.match(r'^(sub|function|property|end|if|then|else|elseif|for|next|to|step|do|loop|while|wend|select|case|exit|on|error|resume|goto|set|let|call|const|dim|as|new|nothing|null|empty|true|false|me|and|or|not|mod|xor|eqv|imp|is|like|redim|preserve|erase|each|in|with|to|byval|byref|optional|paramarray|public|private|friend|static|type|enum|declare|lib|alias|get|put|close|open|print|write|input|output|append|random|binary|lock|unlock|seek|line|width|length|attribute|option|explicit|base|compare|text|binary|option|private|explicit)$', low):
                continue
            problems.append((no, nm, ln.rstrip()))

    if not problems:
        print("Undeclared identifiers: none")
        return 0

    seen = set()
    for no, nm, ln in problems:
        if nm.lower() in seen:
            continue
        seen.add(nm.lower())
        print(f"  line {no}: {nm}   |  {ln.strip()[:90]}")
    return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))