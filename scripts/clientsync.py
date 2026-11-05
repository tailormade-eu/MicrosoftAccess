"""Compare the om* modules of a client project with this library, and copy chosen ones across.

    python -I scripts/clientsync.py report <client modules folder>
    python -I scripts/clientsync.py copy   <client modules folder> <module> [<module> ...]

report   one section per situation, nothing is changed. Exit code 1 when anything needs attention.
copy     puts the library version of the named modules (file names without extension) in the client folder.
         Only the names you give are touched; run `report` first and decide per module.

The hash is the one of scripts/libstamp.py (same normalize), so a module that is only formatted differently,
or has another case of identifiers, counts as identical. See docs/CLIENT-SYNC.md for the working method.

Two comparisons. The stamp hash of libstamp.py is exact about the text (apart from case and trailing spaces) and must stay as it
is: omLibraryStamp.bas calculates the same hash inside Access. This script therefore compares the CODE a second time, with a
looser form (code_form): outside string literals the case is ignored and runs of spaces and tabs count as one, indentation and
blank lines are ignored; inside string literals nothing is touched, so "A" and "a" still differ.

Situations (a module of the client folder whose name starts with om):
  IDENTICAL          same text as the library (after normalizing): nothing to do
  IDENTICAL, NO STAMP  same code, but the client copy has no stamp line. Nothing is wrong now; without the stamp a later library
                     change is reported as NO STAMP, DIFFERS and not as CLIENT BEHIND. `copy` puts the stamped version in
  FORMATTING ONLY    same code, but another case of identifiers, other spaces/indentation or other blank lines.
                     No real difference; `copy` realigns it so the hash matches again
  CLIENT BEHIND      the client copy is untouched (it still matches its own stamp) and the library has moved on: update it
  CHANGED LOCALLY    the client copy no longer matches its own stamp, and differs from the library: someone edited it
                     in the client project. Judge the change: take it into the library, or discard it
  NO STAMP, DIFFERS  an old or foreign copy without a stamp that differs from the library: judge it like CHANGED LOCALLY
  LOCAL ONLY         not in the library: put it in the library, or rename it with the client prefix
"""
import pathlib
import re
import shutil
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import libstamp  # noqa: E402  (same folder)


def lib_modules():
    return {p.stem.lower(): p for p in libstamp.module_files()}


def client_modules(folder):
    return sorted(p for p in pathlib.Path(folder).iterdir()
                  if p.suffix.lower() in ('.bas', '.cls') and p.name.lower().startswith('om'))


def code_form(text):
    """Looser than libstamp.normalize: the code without its formatting (see the module docstring)."""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    out = []
    for line in libstamp.strip_class_header(lines):
        if line.startswith('Attribute ') or line.lstrip(' ').startswith(libstamp.STAMP_PREFIX):
            continue
        res = []
        in_str = False
        i = 0
        while i < len(line):
            ch = line[i]
            if in_str:
                res.append(ch)
                if ch == '"':
                    if i + 1 < len(line) and line[i + 1] == '"':
                        res.append('"')
                        i += 1
                    else:
                        in_str = False
            elif ch == '"':
                in_str = True
                res.append(ch)
            elif ch == "'":
                res.append(' '.join(line[i:].lower().split()))   # a comment: case and spaces do not matter
                break
            elif ch in ' \t':
                if res and res[-1] != ' ':
                    res.append(' ')
            else:
                res.append(ch.lower())
            i += 1
        s = ''.join(res).strip()
        if s:
            out.append(s)
    return '\n'.join(out)


def line_set(text):
    return set(code_form(text).split('\n'))


def classify(client_path, lib_path):
    ctext = libstamp.read(client_path)
    csha = libstamp.sha8(ctext)
    if lib_path is None:
        return 'LOCAL ONLY', 'not in the library', None
    ltext = libstamp.read(lib_path)
    lsha = libstamp.sha8(ltext)
    if code_form(ctext) == code_form(ltext):
        if csha == lsha:
            if libstamp.current_stamp(ctext) is None and libstamp.current_stamp(ltext) is not None:
                # same code, but the client copy has no stamp line: a later library change can then only be reported as
                # NO STAMP, DIFFERS instead of CLIENT BEHIND. `copy` adds the stamp.
                return 'IDENTICAL, NO STAMP', 'same code; the stamp line is missing (copy adds it)', None
            return 'IDENTICAL', '', None
        return 'FORMATTING ONLY', 'same code; case, spaces or blank lines differ', None
    own = libstamp.current_stamp(ctext)
    only_client = len(line_set(ctext) - line_set(ltext))
    only_lib = len(line_set(ltext) - line_set(ctext))
    detail = f'{only_client} lines only in the client, {only_lib} only in the library'
    if only_client == 0 and only_lib == 0:
        detail = 'the same lines in another order or number of times'
    if own is None:
        return 'NO STAMP, DIFFERS', detail, (only_client, only_lib)
    if own[1] == csha:
        return 'CLIENT BEHIND', f'client {own[0]} sha={own[1]}, library sha={lsha}; ' + detail, (only_client, only_lib)
    return 'CHANGED LOCALLY', f'stamp says sha={own[1]}, text hashes to {csha}; ' + detail, (only_client, only_lib)


def class_attributes(text):
    """The Attribute lines of a module (VB_PredeclaredId, VB_Exposed ...) except the name. They change how a class
    behaves (a predeclared class has a default instance), but the code comparison does not look at them."""
    attrs = {}
    for line in text.replace('\r\n', '\n').split('\n'):
        m = re.match(r'^Attribute\s+(VB_\w+)\s*=\s*(.*?)\s*$', line)
        if m and m.group(1).lower() != 'vb_name':
            attrs[m.group(1)] = m.group(2).lower()
    return attrs


def report(folder):
    lib = lib_modules()
    groups = {}
    seen = set()
    attr_diffs = []
    for p in client_modules(folder):
        key = p.stem.lower()
        seen.add(key)
        kind, detail, _ = classify(p, lib.get(key))
        groups.setdefault(kind, []).append((p.stem, detail))
        if key in lib and p.suffix.lower() == '.cls':
            a, b = class_attributes(libstamp.read(p)), class_attributes(libstamp.read(lib[key]))
            if a != b:
                diff = ', '.join(f'{k}: client {a.get(k, "-")} / library {b.get(k, "-")}' for k in sorted(set(a) | set(b)) if a.get(k) != b.get(k))
                attr_diffs.append((p.stem, diff))
    order = ['CHANGED LOCALLY', 'NO STAMP, DIFFERS', 'CLIENT BEHIND', 'LOCAL ONLY', 'IDENTICAL, NO STAMP', 'FORMATTING ONLY', 'IDENTICAL']
    attention = 0
    for kind in order:
        items = groups.get(kind, [])
        if not items:
            continue
        print(f'{kind} ({len(items)})')
        if kind in ('IDENTICAL', 'FORMATTING ONLY', 'IDENTICAL, NO STAMP'):
            print('   ' + ', '.join(n for n, _ in items))
        else:
            attention += len(items)
            for name, detail in items:
                print(f'   {name}' + (f'  - {detail}' if detail else ''))
    if attr_diffs:
        attention += len(attr_diffs)
        print(f'CLASS ATTRIBUTES DIFFER ({len(attr_diffs)})  - the code may be the same, the class does not behave the same')
        for name, diff in attr_diffs:
            print(f'   {name}  - {diff}')
    missing = sorted(lib[k].stem for k in lib if k not in seen)
    print(f'\nlibrary modules the client does not have: {len(missing)}' + (' (not an issue: a client takes what it uses)' if missing else ''))
    print(f'needs attention: {attention}')
    return 1 if attention else 0


def copy(folder, names):
    lib = lib_modules()
    code = 0
    for name in names:
        src = lib.get(name.lower())
        if src is None:
            print(f'{name}: not in the library')
            code = 1
            continue
        dst = pathlib.Path(folder) / src.name
        shutil.copyfile(src, dst)
        print(f'{name}: copied {src.name}')
    return code


def main(argv):
    if len(argv) >= 3 and argv[1] == 'report':
        return report(argv[2])
    if len(argv) >= 4 and argv[1] == 'copy':
        return copy(argv[2], argv[3:])
    print(__doc__)
    return 2


if __name__ == '__main__':
    sys.exit(main(sys.argv))
