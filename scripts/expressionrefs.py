"""List the functions that forms, reports, queries and macros call by name, and that no module defines.

    python -I scripts/expressionrefs.py report <client source folder>

The VBA compiler does not look inside a field expression (`=RoundHalfUp([Total], 2)`), a query or a macro action: a function
that is missing there only fails when the form, report or query is opened ("undefined function", `#Name?`). A project that
compiles can still have such holes, for example after a module was replaced by a newer one that lacks a function.

What is read (the add-in source format of a client project: modules, forms, reports, queries, macros):
  forms/*.form, reports/*.report   every property value that starts with "=", and DefaultValue, ValidationRule, Filter,
                                   OrderBy, RowSource (continuation lines are joined)
  queries/*.sql                    the SQL text
  macros/*.macro                   Argument and Condition values (RunCode "Start()" calls the function Start)

A call is an identifier followed by "(". Not counted: text between quotes, [bracketed names], #dates#, the built-in Access/VBA/SQL
functions (BUILTINS below) and SQL keywords. What is left is looked up in the modules:
  RESOLVED        a Public Function (or one without a modifier) in a standard module (.bas), or any Function (also Private) in the
                  form/report's own module: tested, a Private function in the form's own module works in that form's expressions
  PRIVATE         defined Private in a standard module: not tested here whether Access refuses it, reported as a doubt
  SUB             defined as a Sub: an expression cannot call a Sub
  CLASS           only in a class module: an expression cannot call it
  MISSING         defined nowhere
Exit code 1 when something is not RESOLVED. The built-in list is not complete: a legitimate Access function that is missing
from it shows up as MISSING; add it to BUILTINS.
"""
import pathlib
import re
import sys
from collections import defaultdict

BUILTINS = {n.lower() for n in """
Abs Asc AscB AscW Atn Avg CBool CByte CCur CDate CDbl CDec Choose Chr ChrB ChrW CInt CLng CSng CStr CVar Cos Count CurDir
CurrentDb CurrentUser CVDate Date DateAdd DateDiff DatePart DateSerial DateValue Day DAvg DCount DFirst DLast DLookup DMax DMin
DStDev DStDevP DSum DVar DVarP Environ Eval Exp Fix Format FormatCurrency FormatDateTime FormatNumber FormatPercent
First Forms Hex Hour IIf InStr InStrB InStrRev Int IsArray IsDate IsEmpty IsError IsNull IsNumeric IsObject Join Last LBound LCase
Left LeftB Len LenB Log LTrim Max Mid MidB Min Minute Month MonthName Now Nz Oct Partition Replace Reports Right RightB Rnd Round
RTrim Second Sgn Sin Space Split Sqr StDev StDevP Str StrComp StrConv String StrReverse Sum Switch Tan Time Timer TimeSerial
TimeValue Trim TypeName UBound UCase Val Var VarP VarType Weekday WeekdayName Year Application Screen TempVars Nz
Format$ Left$ Right$ Mid$ Trim$ LTrim$ RTrim$ UCase$ LCase$ Chr$ Str$ Space$ String$ Hex$ Oct$ Date$ Time$
Count Sum Avg Min Max First Last StDev Var
IN AND OR NOT XOR EQV IMP MOD LIKE IS BETWEEN EXISTS VALUES ALL ANY SOME SELECT FROM WHERE JOIN INNER LEFT RIGHT OUTER ON AS BY
GROUP ORDER HAVING UNION DISTINCT DISTINCTROW TOP PERCENT INSERT INTO UPDATE SET DELETE TRANSFORM PIVOT PARAMETERS WITH OWNERACCESS
OPTION CREATE TABLE ALTER ADD COLUMN DROP INDEX PRIMARY KEY CONSTRAINT REFERENCES UNIQUE CHECK FOREIGN CASCADE
VARCHAR CHAR TEXT LONGTEXT MEMO DECIMAL NUMERIC FLOAT DOUBLE INTEGER INT SINGLE BINARY VARBINARY CURRENCY LONG SHORT BYTE
YESNO BIT DATETIME COUNTER AUTOINCREMENT IDENTITY CHARACTER
""".split()}

DEFINE = re.compile(r'^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Function|Sub)\s+(\w+)', re.I)
PROP = re.compile(r'^\s*(\w+)\s*=\s*"(.*)"\s*$')
CONT = re.compile(r'^\s+"(.*)"\s*$')
EXPR_KEYS = {'defaultvalue', 'validationrule', 'filter', 'orderby', 'rowsource', 'controlsource', 'caption'}
CALL = re.compile(r'(?<![\w.!\]])([A-Za-z_]\w*\$?)\s*\(')


def read(path):
    b = path.read_bytes()
    if b[:2] in (b'\xff\xfe', b'\xfe\xff'):
        return b.decode('utf-16')
    try:
        return b.decode('utf-8-sig')
    except UnicodeDecodeError:
        return b.decode('cp1252', 'replace')


def definitions(text):
    """name.lower() -> set of kinds: 'public-function', 'private-function', 'sub'"""
    out = defaultdict(set)
    for line in text.replace('\r\n', '\n').split('\n'):
        m = DEFINE.match(line)
        if not m:
            continue
        private = (m.group(1) or '').strip().lower() == 'private'
        kind = m.group(3).lower()
        out[m.group(4).lower()].add('sub' if kind == 'sub' else ('private-function' if private else 'public-function'))
    return out


def strip_literals(expr):
    expr = re.sub(r'"(?:[^"]|"")*"', ' ', expr)          # string literals
    expr = re.sub(r'\[[^\]]*\]', ' ', expr)               # [bracketed names]
    expr = re.sub(r'#[^#]*#', ' ', expr)                  # #dates#
    return expr


TABLE_WORDS = {'into', 'table', 'from', 'join', 'update', 'index', 'view'}   # a name after these is a table, not a function


def calls(expr):
    s = strip_literals(expr)
    out = set()
    for m in CALL.finditer(s):
        prev = re.search(r'(\w+)\s+$', s[:m.start()])
        if prev and prev.group(1).lower() in TABLE_WORDS:
            continue
        out.add(m.group(1))
    return out


def property_values(text):
    """(key, value) pairs of a form/report text; continuation lines are glued to the value before them."""
    lines = text.replace('\r\n', '\n').split('\n')
    i = 0
    while i < len(lines):
        m = PROP.match(lines[i])
        if m:
            key, val = m.group(1), m.group(2)
            j = i + 1
            while j < len(lines) and CONT.match(lines[j]) and not PROP.match(lines[j]):
                val += CONT.match(lines[j]).group(1)
                j += 1
            yield key, val.replace('""', '"')
            i = j
        else:
            i += 1


def report(root):
    root = pathlib.Path(root)
    std = defaultdict(set)         # defined in a standard module
    cls = defaultdict(set)         # defined in a class module
    for p in (root / 'modules').iterdir() if (root / 'modules').exists() else []:
        if p.suffix.lower() == '.bas':
            for k, v in definitions(read(p)).items():
                std[k] |= v
        elif p.suffix.lower() == '.cls':
            for k, v in definitions(read(p)).items():
                cls[k] |= v
    used = defaultdict(list)       # name -> [(object, own_defs)]

    def note(name, where, own):
        if name.lower() not in BUILTINS:
            used[name.lower()].append((where, own, name))

    for folder, ext in (('forms', '.form'), ('reports', '.report')):
        d = root / folder
        if not d.exists():
            continue
        for p in sorted(d.glob('*' + ext)):
            own = {}
            codefile = p.with_suffix('.cls')
            if codefile.exists():
                own = definitions(read(codefile))
            for key, val in property_values(read(p)):
                if key.lower() == 'rowsource' and not re.match(r'\s*(=|select|transform|insert|update|delete)', val, re.I):
                    continue          # a saved query or a value list, not an expression
                if val.startswith('=') or key.lower() in EXPR_KEYS - {'caption'}:
                    for name in calls(val):
                        note(name, f'{folder}/{p.name}', own)
    d = root / 'queries'
    if d.exists():
        for p in sorted(d.glob('*.sql')):
            for name in calls(read(p)):
                note(name, f'queries/{p.name}', {})
    d = root / 'macros'
    if d.exists():
        for p in sorted(d.glob('*.macro')):
            for key, val in property_values(read(p)):
                if key.lower() in ('argument', 'condition', 'expression'):
                    for name in calls(val):
                        note(name, f'macros/{p.name}', {})

    status = {}
    for key, entries in used.items():
        for where, own, name in entries:
            kinds = std.get(key, set()) | own.get(key, set())
            own_kinds = own.get(key, set())
            if 'public-function' in kinds or 'private-function' in own_kinds:
                # measured: a Private function in the form's/report's OWN module is callable from that object's expressions
                s = 'RESOLVED'
            elif 'private-function' in kinds:
                s = 'PRIVATE'
            elif 'sub' in kinds:
                s = 'SUB'
            elif key in cls:
                s = 'CLASS'
            else:
                s = 'MISSING'
            status.setdefault(key, (s, name, []))[2].append(where)
            if s != 'RESOLVED':
                status[key] = (s, name, status[key][2])
    bad = {k: v for k, v in status.items() if v[0] != 'RESOLVED'}
    resolved = [v[1] for v in status.values() if v[0] == 'RESOLVED']
    print(f'calls found: {len(status)} distinct names | resolved to a module function: {len(resolved)} | not resolved: {len(bad)}')
    for s in ('MISSING', 'PRIVATE', 'SUB', 'CLASS'):
        items = sorted((v for v in bad.values() if v[0] == s), key=lambda v: v[1].lower())
        if items:
            print(f'\n{s} ({len(items)})')
            for _, name, where in items:
                uniq = sorted(set(where))
                print(f'   {name}  - {len(where)} use(s) in {len(uniq)} object(s): ' + ', '.join(uniq[:4]) + (' ...' if len(uniq) > 4 else ''))
    return 1 if bad else 0


def main(argv):
    if len(argv) >= 3 and argv[1] == 'report':
        return report(argv[2])
    print(__doc__)
    return 2


if __name__ == '__main__':
    sys.exit(main(sys.argv))
