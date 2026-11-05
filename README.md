# MicrosoftAccess

This MicrosoftAccess Library is under version control with the [MS Access VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin). The source files live in `src/VCS/MicrosoftAccess.accdb.src`. How the library is built from source, compiled and tested (and how to do the same for a client database) is in [docs/BUILD-AND-TEST.md](docs/BUILD-AND-TEST.md). How to keep a client project in step with the library (drift report, copy, the working method) is in [docs/CLIENT-SYNC.md](docs/CLIENT-SYNC.md).

## Overview

A collection of reusable VBA modules and utilities for Microsoft Access projects. The library provides functions for database connection, string manipulation, table linking, and more.

## Features

- Easy table relinking and connection string management
- String and array utility functions
- Support for ADODB and DAO operations
- Tools for importing/exporting Access objects

## Installation

1. Build `MicrosoftAccess.accdb` from `src/VCS/MicrosoftAccess.accdb.src` with the MS Access VCS Add-in, or take the individual `.bas`/`.cls` files from the `modules` folder.
2. Import the modules you need into your Microsoft Access project via the VBA editor.

## Usage

Example: Relink tables using the library

```vba
' Relink all tables defined in SSMAA_ODBC_Tables
Call LinkUsingSSMA()
```

## Dependencies

- Microsoft Access (tested with version 16 and above)
- References: Microsoft ActiveX Data Objects Library, Microsoft DAO Object Library, Microsoft Scripting Library
- MS Access VCS Add-in for exporting/importing modules

## Folder Structure

- `src/VCS/MicrosoftAccess.accdb.src/modules/` – The modules (`.bas`/`.cls`). **`om*` = the library** (stamped, copied into customer databases); `tst*` = tests and `ex*` = examples live next to them but are not part of the library and have no stamp
- `omSSMAAConnector.bas` – Main module for table linking and connection management
- `omStringFunctions.bas` – String utility functions
- `omBankAccountFunctions.bas` – IBAN validation and formatting utilities
- `databases/` – Example databases
- `tests/` – Sample files that the `tst*` modules need (for example `ocr-sample.png` for `tstTesseractOCR`)
- `scripts/` – `libstamp.py`, the tool that writes the module stamps

## Module stamps

Every `om*` module carries one stamp line, right after the `Option` lines:

```vba
Option Explicit
'@lib 2026-10-10 sha=3fa91c07
```

- `sha` is the first 8 hex characters of a SHA-256 hash of the module text. The date is the day the
  hash last changed (set automatically, no version number to maintain).
- **How the hash is made** (identical in Python and in VBA): take the module text, drop the class
  header (`VERSION`/`BEGIN`/`END`), every `Attribute` line and the stamp line itself; make everything
  lower case (the VBA editor rewrites the case of identifiers, which would otherwise look like a
  change); remove trailing spaces and tabs; remove leading and trailing blank lines; use LF line
  ends; hash the UTF-8 bytes.
- **In the library** run `python -I scripts/libstamp.py apply` before every commit. It adds or
  refreshes the stamps and writes `library-stamps.txt` (`name;date;sha`). `check` lists modules whose
  stamp is missing or out of date (exit code 1), `manifest` only rewrites `library-stamps.txt`.
- **In a customer database** load the library modules (including `omLibraryStamp` and the hash
  modules `omHashAlgorithms`, `omHashAlgorithmsFunctions`) and run:

  ```vba
  Debug.Print omLibraryStamp.LibraryCheck("D:\path	o\library-stamps.txt")
  ```

  It prints one line per module that needs attention: `NO STAMP`, `CHANGED LOCALLY` (the text no
  longer matches its own stamp), `LIBRARY NEWER` (the library has another hash for that module) or
  `LOCAL ONLY` (the module is not in the library). Empty output means everything is in order.
  Without the file name only the local check runs. `GetModuleHash(name)` and `GetModuleStamp(name)`
  give the calculated hash and the stored stamp of one module.

## Module index

Zie [MODULE-INDEX.md](MODULE-INDEX.md) voor een overzicht van alle modules met hun publieke functies/subs. Check deze index (of het bronbestand zelf) vóór je een nieuwe hulpfunctie schrijft — de kans is groot dat ze al bestaat.

## Doel — deze library blijft groeien

Dit is een levende, actief onderhouden library, geen archief. Bedoeling:
- **Verder opbouwen** — nieuwe herbruikbare functies komen hier terecht, niet als losse/dubbele kopie in een klantproject.
- **Committen** — wijzigingen hier horen gecommit te worden, ook als ze vanuit een klantproject-sessie ontstaan (zoals `ClearMultiValueField`, toegevoegd 2026-07-28 vanuit een klantproject).
- **Verbeteren** — bestaande functies mogen gerefactored/uitgebreid worden zodra een project een tekortkoming blootlegt.
- **Index onderhouden** — `MODULE-INDEX.md` moet meebewegen bij elke nieuwe/gewijzigde functie, anders wordt hij snel stale en onbetrouwbaar.

Klantprojecten importeren enkel wat ze nodig hebben (individuele functies of modules, geen automatische volledige clone) — zie de losse module-dependencies vóór je een heel bestand kopieert (bv. `omFormFunctions.bas` hangt af van `omStringFunctions`/`omUserFunctions` voor sommige functies, niet voor allemaal).

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes and improvements.

## License

This project is licensed under the GNU General Public License (GPL). See the LICENSE file for details.

## Contact

For questions or support, contact: [jara@tailormade.eu]
