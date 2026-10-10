# MicrosoftAccess

This MicrosoftAccess Library is under version control with the [MS Access VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin). The source files live in `src/VCS/MicrosoftAccess.accdb.src`.

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

- `src/VCS/MicrosoftAccess.accdb.src/modules/` – The modules (`.bas`/`.cls`), mostly named `om*`
- `omSSMAAConnector.bas` – Main module for table linking and connection management
- `omStringFunctions.bas` – String utility functions
- `omBankAccountFunctions.bas` – IBAN validation and formatting utilities
- `databases/` – Example databases

## Module index

Zie [MODULE-INDEX.md](MODULE-INDEX.md) voor een overzicht van alle modules met hun publieke functies/subs. Check deze index (of het bronbestand zelf) vóór je een nieuwe hulpfunctie schrijft — de kans is groot dat ze al bestaat.

## Doel — deze library blijft groeien

Dit is een levende, actief onderhouden library, geen archief. Bedoeling:
- **Verder opbouwen** — nieuwe herbruikbare functies komen hier terecht, niet als losse/dubbele kopie in een klantproject.
- **Committen** — wijzigingen hier horen gecommit te worden, ook als ze vanuit een klantproject-sessie ontstaan (zoals `ClearMultiValueField`, toegevoegd 2026-07-28 vanuit client Cards).
- **Verbeteren** — bestaande functies mogen gerefactored/uitgebreid worden zodra een project een tekortkoming blootlegt.
- **Index onderhouden** — `MODULE-INDEX.md` moet meebewegen bij elke nieuwe/gewijzigde functie, anders wordt hij snel stale en onbetrouwbaar.

Klantprojecten importeren enkel wat ze nodig hebben (individuele functies of modules, geen automatische volledige clone) — zie de losse module-dependencies vóór je een heel bestand kopieert (bv. `omFormFunctions.bas` hangt af van `omStringFunctions`/`omUserFunctions` voor sommige functies, niet voor allemaal).

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes and improvements.

## License

This project is licensed under the GNU General Public License (GPL). See the LICENSE file for details.

## Contact

For questions or support, contact: [jara@tailormade.eu]
