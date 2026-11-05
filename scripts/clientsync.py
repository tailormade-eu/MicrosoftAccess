"""Compare the om* modules of a client project with this library, and copy chosen ones across.

    python -I scripts/clientsync.py report <client modules folder>
    python -I scripts/clientsync.py copy   <client modules folder> <module> [<module> ...]

report   one section per situation, nothing is changed. Exit code 1 when anything needs attention.
copy     puts the library version of the named modules (file names without extension) in the client folder.
         Only the names you give are touched; run `report` first and decide per module.

The hash is the one of scripts/libstamp.py (same normalize), so a module that is only formatted differently,
or has another case of identifiers, counts as identical. See docs/CLIENT-SYNC.md for the working method.

Situations (a module of the client folder whose name starts with om):
  IDENTICAL          same text as the library (after normalizing): nothing to do
  CLIENT BEHIND      the client copy is untouched (it still matches its own stamp) and the library has moved on: update it
  CHANGED LOCALLY    the client copy no longer matches its own stamp, and differs from the library: someone edited it
                     in the client project. Judge the change: take it into the library, or discard it
  NO STAMP, DIFFERS  an old or foreign copy without a stamp that differs from the library: judge it like CHANGED LOCALLY
  LOCAL ONLY         not in the library: put it in the library, or rename it with the client prefix
"""
import pathlib
import shutil
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import libstamp  # noqa: E402  (same folder)


def lib_modules():
    return {p.stem.lower(): p for p in libstamp.module_files()}


def client_modules(folder):
    return sorted(p for p in pathlib.Path(folder).iterdir()
                  if p.suffix.lower() in ('.bas', '.cls') and p.name.lower().startswith('om'))


def line_set(text):
    return set(libstamp.normalize(text).split('\n'))


def classify(client_path, lib_path):
    ctext = libstamp.read(client_path)
    csha = libstamp.sha8(ctext)
    if lib_path is None:
        return 'LOCAL ONLY', 'not in the library', None
    ltext = libstamp.read(lib_path)
    lsha = libstamp.sha8(ltext)
    if csha == lsha:
        return 'IDENTICAL', '', None
    own = libstamp.current_stamp(ctext)
    only_client = len(line_set(ctext) - line_set(ltext))
    only_lib = len(line_set(ltext) - line_set(ctext))
    detail = f'{only_client} lines only in the client, {only_lib} only in the library'
    if only_client == 0 and only_lib == 0:
        detail = 'only blank lines or the order of lines differ'
    if own is None:
        return 'NO STAMP, DIFFERS', detail, (only_client, only_lib)
    if own[1] == csha:
        return 'CLIENT BEHIND', f'client {own[0]} sha={own[1]}, library sha={lsha}; ' + detail, (only_client, only_lib)
    return 'CHANGED LOCALLY', f'stamp says sha={own[1]}, text hashes to {csha}; ' + detail, (only_client, only_lib)


def report(folder):
    lib = lib_modules()
    groups = {}
    seen = set()
    for p in client_modules(folder):
        key = p.stem.lower()
        seen.add(key)
        kind, detail, _ = classify(p, lib.get(key))
        groups.setdefault(kind, []).append((p.stem, detail))
    order = ['CHANGED LOCALLY', 'NO STAMP, DIFFERS', 'CLIENT BEHIND', 'LOCAL ONLY', 'IDENTICAL']
    attention = 0
    for kind in order:
        items = groups.get(kind, [])
        if not items:
            continue
        print(f'{kind} ({len(items)})')
        if kind == 'IDENTICAL':
            print('   ' + ', '.join(n for n, _ in items))
        else:
            attention += len(items)
            for name, detail in items:
                print(f'   {name}' + (f'  - {detail}' if detail else ''))
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
