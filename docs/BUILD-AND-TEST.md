# Building, compiling and testing

How this library is built from source, checked and tested, and how the VCS add-in is used for it. Everything here was
done on Windows 11 with 64-bit Access 365 and the [MSAccessVCS add-in](https://github.com/joyfullservice/msaccess-vcs-addin)
5.1.0. The scripts are in [`tools/vcs`](../tools/vcs) and [`tools/test`](../tools/test).

## 1. The source is the truth

`src/VCS/MicrosoftAccess.accdb.src` is an export of the library database in the add-in's format (one file per object:
`modules/*.bas|cls`, `forms/*.bas|cls`, `tbldefs`, `queries`, `macros`, plus `vcs-options.json`, `proj-properties.json`,
`vbe-references.json`, ...). The `.accdb` is a build product, not a source.

| What | How |
|---|---|
| Source to database | `Build` of the add-in (ribbon: *Build From Source*; or `API.Build(sourceFolder)` over COM) |
| Database to source | `Export` of the add-in (ribbon: *Export Source*) |
| Edit a module | edit the file in `modules/`, then build (or import that object) |
| Module stamp | `python -I scripts/libstamp.py apply` (see README, *Module stamps*); the Excel modules: add `excel` |

Settings that matter, in `vcs-options.json` and `proj-properties.json`:

- **`BreakOnError`** (`true` stops the VBE on every error): fine on your desk, fatal in a hidden Access, where nobody sees the Stop.
- **`UseEnvForConnections`**: linked tables carry `env:conn_x` instead of a connection string. Without that environment
  variable Access asks for a data source for every linked table.
- **`VCS Build Path`**: where *this* project builds to. A build that finds another database name opens
  *Build Name Conflict*, and **OK builds over the file named in the source (your real database)**. Never press OK in a test.

## 2. Build a copy: `tools/vcs/Build-FromSource.ps1`

```powershell
.\tools\vcs\Build-FromSource.ps1 -SourceFolder src\VCS\MicrosoftAccess.accdb.src -DatabaseName MicrosoftAccess.accdb
# a client project that cannot reach its SQL Server from here:
.\tools\vcs\Build-FromSource.ps1 -SourceFolder D:\repos\App\App.mdb -DatabaseName App.mdb -DropLinkedTables -DropObjects reports/Broken
```

It builds in a copy (`%TEMP%\vcs-build\<name>`), points *VCS Build Path* at that copy, turns `BreakOnError` off in the copy,
presses Cancel on a *Build Name Conflict*, and stops only the Access it started. Read the lines `ERROR` and `CRITICAL` in the
log it prints; `Error 48 (Error in loading DLL)` while importing *vbe references* means a referenced library is not installed on
this pc. The build goes on, but code that uses a type from that library will not compile.

If the log shows *Error getting theme file ... Office Theme.thmx* with the *Compressed (zipped) Folders Error* window, the add-in could not
zip the theme folder in that work folder (seen once in `%TEMP%`, not in another folder; cause unknown). Add `-DropThemes`: the build then
ends clean in 13 s, without the theme.

## 3. Compile check

A database that builds is not a database that compiles. Compile it:

1. keep the database open in a hidden Access: `Start-Process pwsh -ArgumentList '-NoProfile','-File',tools\vcs\Hold-Database.ps1,<db> -WindowStyle Hidden` and wait 30 s
2. call the MCP tool `vcs_compile_vba(<db>)` (msaccess-vcs-mcp). It returns `success` true/false, no line number
3. `New-Item <folder>\stop.flag` to let the script close Access

On `false`, open the VBE (*Debug > Compile*): Access jumps to the line. Do the check on a deliberately broken module once, to see
it fail (`RunCommand(126)` does *not* report a compile error in a hidden Access, so it cannot be used).

## 4. Tests: a hidden Access, text out

[`tools/test/Test-AccessTemplate.ps1`](../tools/test/Test-AccessTemplate.ps1) is the pattern: copy a built database, open it
hidden, load a small helper module with `LoadFromText`, call it with `$acc.Run`, compare text, quit. For Excel:
[`Test-ExcelTemplate.ps1`](../tools/test/Test-ExcelTemplate.ps1). Rules that save time:

- **One line per check**, `OK name` or `FOUT name expected [..] got [..]`, so a run reads in two seconds.
- **Prove the test can fail** once before you trust it.
- **Test the thing, not a copy of it**: load the module under test from the working tree.
- **A wait with a timeout, and on timeout read the windows** of your test Access (UI Automation). A modal window is the usual hang.
- **Stop only your own Access** (command line contains `-Embedding`), and check with `Get-Process` that it is gone; a
  "stopped" message from a tool stops the shell, not the process.
- Tables and forms a function needs are created by the function (`omPrinters`, `omFiles`, `omHelpFile`, `Calendar`), so a test
  starts from an empty database.

### Traps we hit (symptom, cause)

| Symptom | Cause |
|---|---|
| `Value cannot be null (Parameter 'ptr')` from `Run` | the VBA function returned an empty string; return `[empty]` |
| a test says an object is `Nothing` when it was set to `Nothing` | `Dim x As New T` re-creates `x` on the next use; use `Set x = New T` in test code |
| Access waits forever after `CreateReport` | a new report must be `DoCmd.Save`d before `Close` |
| garbage characters, or `VERSION 1.0 CLASS` as code | `LoadFromText` reads Windows-1252; a `.cls` cannot be loaded this way |
| `RPC server unavailable (0x800706BE)` when touching one form | a control on that form crashes Access through COM; skip the form |
| Excel: *version 1.00 class* when filling a class | strip the `VERSION ... END` header and `Attribute` lines first |
| Excel: `Module1` appears next to the module you replaced | a running macro keeps the old one; replace the code, do not remove and import |
| `Workbooks.Open(f, Format:=xlCSV)` fails | `Format` is the delimiter argument; leave it out |
| `robocopy` "failed" in a script | exit codes 0 to 7 are success |

## 5. Compare a live database with the source

Opening the live file runs its startup form and AutoExec (a connection error window, a login form). On a **copy**:

```powershell
.\tools\vcs\Export-DatabaseText.ps1 -Db C:\temp\copy.mdb -Out C:\temp\live -What code -AllowBypass
.\tools\vcs\Export-DatabaseText.ps1 -Db C:\temp\copy.mdb -Out C:\temp\live -What forms -Skip BadForm
.\tools\vcs\Export-DatabaseText.ps1 -Db C:\temp\copy.mdb -Out C:\temp\live -What reports
```

It holds Shift while opening (bypass), writes `SaveAsText` files, and the library's own normalisation (`scripts/libstamp.py`)
is the right way to compare code (header, attributes and case ignored).

## 6. Changing a client database without the add-in

When the add-in cannot build a client project here (missing references, a broken export), apply the change to a **copy of the real
database** and compile that: delete the old modules, `LoadFromText` the new ones, edit the others in place, compile. The copy
has its references already. Do not open the real file.

## 7. Updating the add-in

The add-in's own repo documents it ([AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md)): edit the
source, then `vcs_rebuild_addin("<repo>\Version Control.accda.src")`. It rebuilds the development copy and installs it under
`%AppData%\MSAccessVCS`; the previous installed file is not kept, so note the version first. 5.1.0 fixed the zip error on the theme
folder (*Compressed (zipped) Folders Error: File not found or no read permission*) that earlier builds showed.

## 8. Memory and processes

Several Access, Excel and .NET compiler processes at once can fill the memory (a tool may stop a run for that). Run one build at a
time, and look at what is still running afterwards: `Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'"`.
