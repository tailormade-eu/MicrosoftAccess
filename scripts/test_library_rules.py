"""Rules for the library modules:  python -I scripts/test_library_rules.py   (from the repo root)

A setting of a client never lives in an om* module as a Const: the next copy of the library overwrites it, and the client
then runs with another value (that is how a client got another login screen after a refresh). The standard value is in
omGlobalDeclarations.ClientSettings_Load, the client changes it in its own ClientSettings_LoadOverride.
"""
import pathlib
import re
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import libstamp  # noqa: E402

BOOL_CONST = re.compile(r"^\s*(Public\s+)?Const\s+(\w+)(\s+As\s+Boolean)?\s*=\s*(True|False)\b", re.I)


class LibraryRulesTests(unittest.TestCase):
    def test_no_boolean_const_in_a_library_module(self):
        found = []
        for p in libstamp.module_files():
            for n, line in enumerate(libstamp.read(p).splitlines(), 1):
                m = BOOL_CONST.match(line)
                if m and not line.lstrip().startswith("'"):
                    found.append(f"{p.name}:{n} {m.group(2)}")
        self.assertEqual(found, [], "a True/False Const in an om* module is a setting of a client: put the standard value in "
                                    "omGlobalDeclarations.ClientSettings_Load and let the client override it")

    def test_the_override_is_never_defined_in_the_library(self):
        found = []
        for p in libstamp.module_files():
            if re.search(r"^\s*(Public\s+)?Sub\s+ClientSettings_LoadOverride\b", libstamp.read(p), re.I | re.M):
                found.append(p.name)
        self.assertEqual(found, [], "ClientSettings_LoadOverride belongs to the client project (a stub here would overwrite it)")

    def test_modules_that_read_a_setting_ask_for_it_first(self):
        text = {p.stem: libstamp.read(p) for p in libstamp.module_files()}
        for module, needle in (("omUserFunctions", "LoginMethodIsSet"), ("omDefaultFunctions", "gDevelopmentMode"),
                               ("omTranslationEngineFunctions", "CurrentLanguageId"), ("omSeqFunctions", "SeqServerUrl")):
            self.assertIn("EnsureClientSettings", text[module], f"{module} reads {needle} but never calls EnsureClientSettings")


if __name__ == "__main__":
    unittest.main()
