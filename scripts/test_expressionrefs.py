"""Tests for expressionrefs:  python -I scripts/test_expressionrefs.py   (from the repo root)

Builds a tiny project in a temporary folder and checks which calls are reported.
"""
import contextlib
import io
import pathlib
import sys
import tempfile
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import expressionrefs  # noqa: E402

MODULE = '''Attribute VB_Name = "modX"
Option Explicit
Public Function PublicOne() As Long
End Function
Function NoModifier() As Long
End Function
Private Function HiddenOne() As Long
End Function
Public Sub OnlyASub()
End Sub
'''
CLASSMOD = '''Option Explicit
Public Function InClassOnly() As Long
End Function
'''
FORM_CODE = '''Option Explicit
Public Function OwnFormFunction() As Long
End Function
'''
REPORT = '''Version =20
Begin Report
    Begin
        Begin TextBox
            ControlSource ="=PublicOne(1)+NoModifier()+Nz([x],0)+Round(2.5)"
            Name ="a"
        End
        Begin TextBox
            ControlSource ="=HiddenOne()+OnlyASub()+InClassOnly()+NoSuchFunction([x])"
            Name ="b"
        End
        Begin TextBox
            ControlSource ="=IIf([t]=""MissingInQuotes()"",1,2)"
            Name ="c"
        End
        Begin TextBox
            ControlSource ="=Public"
                "One(1)+Split"
                "Joined(2)"
            Name ="d"
        End
        Begin TextBox
            ControlSource ="=OwnFormFunction()"
            Name ="e"
        End
    End
End
'''
QUERY = 'SELECT DLookup("a", "t"), MissingInSql([x]) FROM t INSERT INTO Tbl (a) VALUES (1);'
MACRO = '''Begin
    Action ="RunCode"
    Argument ="MissingFromMacro()"
End
Begin
    Action ="RunCode"
    Argument ="PublicOne()"
End
'''


def run(root):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = expressionrefs.report(root)
    return code, buf.getvalue()


class ExpressionRefsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        r = pathlib.Path(self.tmp.name)
        for d in ('modules', 'reports', 'queries', 'macros'):
            (r / d).mkdir()
        (r / 'modules' / 'modX.bas').write_text(MODULE, encoding='utf-8')
        (r / 'modules' / 'clsY.cls').write_text(CLASSMOD, encoding='utf-8')
        (r / 'reports' / 'R.report').write_text(REPORT, encoding='utf-8')
        (r / 'reports' / 'R.cls').write_text(FORM_CODE, encoding='utf-8')
        (r / 'queries' / 'Q.sql').write_text(QUERY, encoding='utf-8')
        (r / 'macros' / 'M.macro').write_text(MACRO, encoding='utf-8')
        self.root = r

    def tearDown(self):
        self.tmp.cleanup()

    def test_reports_what_cannot_be_called(self):
        code, out = run(self.root)
        self.assertEqual(code, 1)
        section = None
        found = {}
        for line in out.splitlines():
            if line and not line.startswith(' ') and '(' in line and line.split(' ')[0].isupper():
                section = line.split(' ')[0]
            elif line.startswith('   ') and section:
                found[line.split()[0]] = section
        expected = {'NoSuchFunction': 'MISSING', 'MissingInSql': 'MISSING', 'MissingFromMacro': 'MISSING',
                    'HiddenOne': 'PRIVATE', 'OnlyASub': 'SUB', 'InClassOnly': 'CLASS'}
        for name, label in expected.items():
            self.assertEqual(found.get(name), label, name)

    def test_does_not_report_what_is_fine(self):
        _, out = run(self.root)
        for name in ('PublicOne', 'NoModifier', 'DLookup', 'Nz', 'IIf', 'OwnFormFunction', 'Tbl', 'MissingInQuotes'):
            self.assertNotIn(name + ' ', out.replace('(', ' '), name)

    def test_continuation_lines_are_joined(self):
        # "=Public" + "One(1)+Split" + "Joined(2)" is PublicOne(1)+SplitJoined(2): the second one is missing
        _, out = run(self.root)
        self.assertIn('SplitJoined', out)

    def test_clean_project_gives_exit_code_0(self):
        (self.root / 'queries' / 'Q.sql').write_text('SELECT PublicOne(1) FROM t;', encoding='utf-8')
        (self.root / 'macros' / 'M.macro').write_text('Begin\n    Action ="RunCode"\n    Argument ="PublicOne()"\nEnd\n', encoding='utf-8')
        (self.root / 'reports' / 'R.report').write_text('Version =20\nBegin Report\n    Begin\n        Begin TextBox\n            ControlSource ="=PublicOne(1)+Nz([x],0)"\n        End\n    End\nEnd\n', encoding='utf-8')
        code, _ = run(self.root)
        self.assertEqual(code, 0)


if __name__ == '__main__':
    unittest.main()
