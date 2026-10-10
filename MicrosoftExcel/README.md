# MicrosoftExcel

Reusable VBA modules for Excel workbooks. The Access library is in `../src/VCS/MicrosoftAccess.accdb.src`; this one is separate because
Access modules start with `Option Compare Database`, which Excel does not accept.

Each module is a plain `.bas` file in `src/modules` (ANSI-safe, CRLF, no BOM): import it with the VBA editor (File > Import File) or
with `omExcelVersionControl.ImportVBA`. Every module carries a stamp line `'@lib <date> sha=<8 hex>` that works as in the Access library
(`python -I scripts/libstamp.py apply excel`, manifest in `library-stamps.txt`).

| Module | What it does | Needs |
|---|---|---|
| `omExcelApplicationFunctions` | Application settings off/on for fast macros, `SheetExists` | |
| `omOnedriveFunctions` | The real local folder of a workbook in OneDrive/SharePoint | |
| `omExcelVBAFunctions` | `Nz`, local number text to number, local formula to VBA notation | |
| `omExcelRecordsetFunctions` | A sheet range of a closed .xlsx as an ADO recordset | ADODB reference, ACE provider |
| `omListObjectFunctions` | Tables: find a row, sort, max, recordset | `omOnedriveFunctions`, `omExcelRecordsetFunctions`, ADODB |
| `omWorksheetFunctions` | Show/hide sheets, copy a template row | |
| `omLockingFunctions` | A `.lock` file next to a project file (who works in it) | |
| `omExcelVersionControl` | Sheets (csv), VBA and sheet code to a `Src` folder for git, and back | `omOnedriveFunctions`, `omExcelApplicationFunctions`, trust access to the VBA project object model |

Tested in a hidden Excel through COM (the test scripts are not part of the library). Not carried over from the projects it came from:
the SHA256 class (`omHashAlgorithms` in the Access library does it in UTF-8) and the SBD-only data modules.
