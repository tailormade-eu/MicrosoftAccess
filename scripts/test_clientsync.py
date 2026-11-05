"""Tests for clientsync.classify:  python -I -m unittest scripts/test_clientsync.py   (from the repo root)

Each case takes a real library module, changes it in one way and checks which situation the script reports.
"""
import pathlib
import re
import sys
import tempfile
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import clientsync  # noqa: E402
import libstamp  # noqa: E402

LIB = {p.stem: p for p in libstamp.module_files()}
NAME = 'omStringFunctions'


def kind_for(text, name=NAME):
    with tempfile.TemporaryDirectory() as d:
        p = pathlib.Path(d) / (name + '.bas')
        p.write_bytes(text.encode('utf-8'))
        return clientsync.classify(p, LIB[name])[0]


class ClientSyncTests(unittest.TestCase):
    def setUp(self):
        self.text = libstamp.read(LIB[NAME])

    def test_same_file_is_identical(self):
        self.assertEqual(kind_for(self.text), 'IDENTICAL')

    def test_blank_lines_and_indentation_are_formatting_only(self):
        changed = self.text.replace('\r\n', '\n').replace('Option Explicit\n', 'Option Explicit\n\n\n', 1)
        changed = re.sub(r'^    ', '        ', changed, count=5, flags=re.M)
        self.assertEqual(kind_for(changed), 'FORMATTING ONLY')

    def test_case_of_identifiers_and_extra_spaces_are_formatting_only(self):
        changed = self.text.replace('Public Function', 'PUBLIC  FUNCTION', 3)
        self.assertEqual(kind_for(changed), 'FORMATTING ONLY')

    def test_case_inside_a_string_is_a_real_difference(self):
        # the first library module that has a string literal in a code line (not a comment, where case does not matter)
        pattern = re.compile(r"^(?!Attribute)[^'\n]*?(\"[A-Za-z][A-Za-z ]{3,}\")", re.M)
        for name, path in sorted(LIB.items()):
            text = libstamp.read(path)
            m = pattern.search(text)
            if m:
                literal = m.group(1)
                changed = text.replace(literal, literal.swapcase(), 1)
                self.assertNotEqual(changed, text)
                self.assertNotIn(kind_for(changed, name), ('IDENTICAL', 'FORMATTING ONLY'), name)
                return
        self.fail('no library module with a string literal found')

    def test_a_real_code_change_is_not_formatting(self):
        changed = self.text + '\nPublic Sub ExtraProcedure()\nEnd Sub\n'
        self.assertNotIn(kind_for(changed), ('IDENTICAL', 'FORMATTING ONLY'))


if __name__ == '__main__':
    unittest.main()
