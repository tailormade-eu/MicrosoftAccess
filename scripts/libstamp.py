"""Stamp every om* module with a content hash: a line  '@lib <date> sha=<8 hex>  after the Option lines.

    python -I scripts/libstamp.py check      list modules whose stamp is missing or out of date
    python -I scripts/libstamp.py apply      write/refresh the stamps (the date changes only when the hash changes)
    python -I scripts/libstamp.py manifest   write library-stamps.txt (name;date;sha) from the stamps in the modules

The hash must be identical to the one omLibraryStamp.bas calculates inside Access, so both sides
normalize the module text the same way (see normalize): no VERSION/BEGIN/END class header, no
"Attribute" lines, no stamp line, lower case (the VBA editor rewrites the case of identifiers),
no trailing spaces or tabs, no leading or trailing blank lines, LF line ends, hash over UTF-8.
"""
import hashlib
import pathlib
import re
import sys
import datetime

ROOT = pathlib.Path(__file__).resolve().parent.parent
MODULES = ROOT / 'src' / 'VCS' / 'MicrosoftAccess.accdb.src' / 'modules'
MANIFEST = ROOT / 'library-stamps.txt'
STAMP = re.compile(r"^'@lib (\d{4}-\d{2}-\d{2}) sha=([0-9a-f]{8})\s*$")
STAMP_PREFIX = "'@lib "


def strip_class_header(lines):
    j = 0
    while j < len(lines) and lines[j].strip() == '':
        j += 1
    if j < len(lines) and lines[j].startswith('VERSION '):
        while j < len(lines) and lines[j].strip() != 'END':
            j += 1
        return lines[j + 1:]
    return lines


def normalize(text):
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    out = []
    for line in strip_class_header(lines):
        if line.startswith('Attribute '):
            continue
        if line.lstrip(' ').startswith(STAMP_PREFIX):
            continue
        out.append(line.lower().rstrip(' \t'))
    return '\n'.join(out).strip('\n')


def sha8(text):
    return hashlib.sha256(normalize(text).encode('utf-8')).hexdigest()[:8]


def module_files():
    return sorted(p for p in MODULES.iterdir() if p.suffix.lower() in ('.bas', '.cls') and p.name.lower().startswith('om'))


def read(path):
    return path.read_bytes().decode('utf-8-sig')


def current_stamp(text):
    for line in text.replace('\r\n', '\n').split('\n'):
        m = STAMP.match(line)
        if m:
            return m.group(1), m.group(2)
    return None


def insert_position(lines):
    """Index where the stamp line goes: after the header, the Attribute lines and the Option lines."""
    i = 0
    if lines and lines[0].startswith('VERSION '):
        while i < len(lines) and lines[i].strip() != 'END':
            i += 1
        i += 1
    while i < len(lines) and lines[i].startswith('Attribute '):
        i += 1
    while i < len(lines) and lines[i].startswith('Option '):
        i += 1
    return i


def stamped_text(text, today):
    """The text with a fresh stamp; unchanged when the stamp is already right."""
    sha = sha8(text)
    cur = current_stamp(text)
    if cur and cur[1] == sha:
        return text
    date = today
    lines = text.replace('\r\n', '\n').split('\n')
    lines = [l for l in lines if not STAMP.match(l)]
    lines.insert(insert_position(lines), "%s%s sha=%s" % (STAMP_PREFIX, date, sha))
    return '\r\n'.join(lines)


def main(argv):
    mode = argv[1] if len(argv) > 1 else 'check'
    today = datetime.date.today().isoformat()
    bad = 0
    manifest = []
    for path in module_files():
        text = read(path)
        sha = sha8(text)
        cur = current_stamp(text)
        if mode == 'apply':
            new = stamped_text(text, today)
            if new != text:
                path.write_bytes(b'\xef\xbb\xbf' + new.encode('utf-8'))
                cur = current_stamp(new)
                print('stamped', path.name, cur[0], cur[1])
        elif mode == 'check':
            if cur is None:
                print('NO STAMP ', path.name)
                bad += 1
            elif cur[1] != sha:
                print('OUT OF DATE', path.name, 'stamp', cur[1], 'actual', sha)
                bad += 1
        if cur:
            manifest.append('%s;%s;%s' % (path.stem, cur[0], cur[1]))
    if mode in ('apply', 'manifest'):
        MANIFEST.write_bytes(('\r\n'.join(manifest) + '\r\n').encode('utf-8'))
        print('manifest', len(manifest), 'modules')
    if mode == 'check':
        print('ok' if bad == 0 else '%d module(s) need a stamp' % bad)
    return 1 if (mode == 'check' and bad) else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
