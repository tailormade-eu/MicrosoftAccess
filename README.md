# MicrosoftAccess

This MicrosoftAccess Library is exported using Oasis-SVN, which can be downloaded from [https://dev2dev.de](https://dev2dev.de).

## Overview

A collection of reusable VBA modules and utilities for Microsoft Access projects. The library provides functions for database connection, string manipulation, table linking, and more.

## Features

- Easy table relinking and connection string management
- String and array utility functions
- Support for ADODB and DAO operations
- Tools for importing/exporting Access objects

## Installation

1. Download or export the modules using Oasis-SVN.
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
- Oasis-SVN for exporting/importing modules

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes and improvements.

## License

This project is licensed under the GNU General Public License (GPL). See the LICENSE file for details.

## Contact

For questions or support, contact: [jara@tailormade.eu]
