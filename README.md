# MicrosoftAccess

This MicrosoftAccess Library is exported using Oasis-SVN, which can be downloaded from [https://dev2dev.de](https://dev2dev.de).

Additionally, this library now supports version control using the [MS Access VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin).

## Overview

A collection of reusable VBA modules and utilities for Microsoft Access projects. The library provides functions for database connection, string manipulation, table linking, and more.

## Features

- Easy table relinking and connection string management
- String and array utility functions
- Support for ADODB and DAO operations
- Tools for importing/exporting Access objects

## Installation

1. Download or export the modules using Oasis-SVN or the MS Access VCS Add-in.
2. Import the `.def` files into your Microsoft Access project via the VBA editor.

## Usage

Example: Relink tables using the library

```vba
' Relink all tables defined in SSMAA_ODBC_Tables
Call LinkUsingSSMA()
```

## Dependencies

- Microsoft Access (tested with version 16 and above)
- References: Microsoft ActiveX Data Objects Library, Microsoft DAO Object Library, Microsoft Scripting Library
- Oasis-SVN or MS Access VCS Add-in for exporting/importing modules

## Folder Structure

- `src/` – Contains source files formatted for both Oasis-SVN and MS Access VCS Add-in.
- `M_omSSMAAConnector.def` – Main module for table linking and connection management
- `M_omStringFunctions.def` – String utility functions
- `M_omBankAccountFunctions.def` – IBAN validation and formatting utilities
- Other `.def` files – Additional utilities and helpers

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes and improvements.

## License

This project is licensed under the GNU General Public License (GPL). See the LICENSE file for details.

## Contact

For questions or support, contact: [jara@tailormade.eu]
