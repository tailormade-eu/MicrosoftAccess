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

The file that is built comes from the source, not from the name you give the script:

- `project.json` `FileFormat`: 10 = Access 2002 `.mdb`, 12 = `.accdb`. A format-2002 database refuses some modern forms (see the traps table).
- `vbe-project.json` `FileName` and `dbs-properties.json` `Name`: the file name, with its extension. The add-in builds to
  `<VCS Build Path>\<that name>`. A script that is told another extension ends up with an empty database of that name next
  to the real build.
- `vbe-references.json`: the references (library GUID and version) that are added to the VBA project during the build.

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

To see which reference fails, add the entries of `vbe-references.json` one by one to an empty database
(`$acc.References.AddFromGuid(guid, major, minor)` in a hidden Access, one `try` per entry). Seen on 64-bit Access: only
*Microsoft DAO 3.6* (`{00025E01-0000-0000-C000-000000000046}`, 5.0) fails, with *Error in loading DLL*, because `dao360.dll`
exists only as a 32-bit file and cannot be installed for 64-bit Office. `stdole` answers *Name conflicts with existing module*,
which is harmless (Access has it already). The fix for a client project is the ACE DAO in `vbe-references.json`
(`{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}`, 12.0): the same object model, available in 32 and 64 bit.

If the log shows *Error getting theme file ... Office Theme.thmx* with the *Compressed (zipped) Folders Error* window, the add-in could not
zip the theme folder in that work folder (seen once in `%TEMP%`, not in another folder; cause unknown). Add `-DropThemes`: the build then
ends clean in 13 s, without the theme.

## 3. Compile check

A database that builds is not a database that compiles. Two ways, on a **copy**:

**3a. In the VBE, with the error position** - [`tools/vcs/Compile-InVbe.ps1`](../tools/vcs/Compile-InVbe.ps1) `-Database <copy>`.
It opens the copy hidden with Shift held (the startup form, for example a login, does not run), runs *Debug > Compile*, presses
OK on the error box and prints `ERRORPOS module=... line=...` with the lines around it. No error: the Compile menu item is greyed
out afterwards (`enabled after: False`) and no position is printed. This is the method to use for a client database.

**3b. With the MCP tool** - for a database without a startup form:

1. keep the database open in a hidden Access: `Start-Process pwsh -ArgumentList '-NoProfile','-File',toolscs\Hold-Database.ps1,'-Database',<db> -WindowStyle Hidden` and wait 30 s
2. call the MCP tool `vcs_compile_vba(<db>)` (msaccess-vcs-mcp). It returns `success` true/false, no line number
3. `New-Item <folder>\stop.flag` to let the script close Access

A startup form (a login window) blocks 3b: the hold script never reaches the end and the tool answers `false` without a reason.
Do the check on a deliberately broken module once, to see it fail (`RunCommand(126)` does *not* report a compile error in a hidden
Access, so it cannot be used).

**DAO 3.6.** A client `.mdb` may still reference *Microsoft DAO 3.6* instead of the ACE DAO. Types that exist only in the ACE
DAO (`DAO.Recordset2`, `Field2`) then give *User-defined type not defined*. Library code that has to run in both declares them
`As Object`.

**3c. A client project with linked tables (no server on this pc).** Build and compile it in this order:

1. copy the source folder; in the copy set `project.json` `FileFormat` and the file name in `vbe-project.json` / `dbs-properties.json`
   if you want an `.accdb` (see the trap table: a format-2002 `.mdb` refuses some forms)
2. `tools\vcs\New-StandInTables.ps1 -Template <copy of the real database> -SourceFolder <the source copy>` writes one
   `tbldefs\<name>.xml` per linked table: a local table with the same fields
3. `tools\vcs\Build-FromSource.ps1 -SourceFolder <the source copy> -DatabaseName X.accdb -DropLinkedTables`
4. `tools\vcs\Compile-InVbe.ps1 -Database <the built file>`

**Tables must exist before the forms and reports are read in.** A form read in without its table has an empty field list, and
`Me.SomeField` in its code does not compile, although the code is fine. Creating the stand-in tables *after* the build, or
re-saving the forms afterwards, did not help (tried). What the compile then reports is real: for one client project it found one
dead event handler (`CostPrice_AfterUpdate`) that referred to a control and field that no longer exist; in the real database
it compiled only because the form kept the old type information.

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
| the script waits and the windows show a *Login* form | the startup form of the client database: open with Shift held (`AllowBypassKey` in the copy), see `Compile-InVbe.ps1` |
| a form is imported but Access will not save it: *changes that are incompatible with the current database format* (error 13122, add-in message *Error while initializing form*) | the layout-grid system of Access 2007 and later, in a project of format 2002 (`project.json` `FileFormat` 10). Two things in the form text, and both must go: the layout properties on the controls (`LayoutGroup`, `GroupTable`, `ColumnStart`/`ColumnEnd`, `RowStart`/`RowEnd`, the four paddings) and any real `EmptyCell` control (an empty cell of a layout; the `EmptyCell` block in the defaults at the top is harmless). Removing only the properties, or only the cell, still fails. Found with `tools/vcs/Shrink-FormText.ps1`; in a format-12 `.accdb` the same forms save. The form is imported even when the save fails, but it is unreliable (no controls in design view, or not exportable) |
| `-Db` is refused as a parameter name | it is an alias of PowerShell's `-Debug`; use `-Database` |
| PowerShell: `DoCmd.OpenForm($n, 0, '', '', '', 3)` fails with *Cannot convert the "" value of type string to type Object* | pass nothing for an optional COM argument, or `[Type]::Missing`; give only the arguments you need (`OpenForm($n, 1)` for design view) |
| a test that opens a form in normal view never ends | normal view runs `Form_Open` and the startup logic (a login box, a message): open in design view (`acDesign` = 1) or read the object with `SaveAsText` |
| `$vbe.VBProjects(1).VBComponents.Item(x).CodeModule.InsertText(...)` fails: *does not contain a method named InsertText* | `InsertText` belongs to the Access `Module` object (`$acc.Modules.Item(x)`); on a VBE `CodeModule` use `AddFromString`. Mind that `DeleteLines` had already run: a failed insert leaves an EMPTY module, and an empty module compiles |
| *Option Compare Database* twice after adding a class module through `VBComponents.Add(2)` | Access puts `Option Compare Database` in a new module itself; do not add it again (compile error on line 3) |
| `Me.SomeField` does not compile in a form's code, although the field exists | the form was imported before its table existed, so its field list is empty. Re-saving it afterwards does not help (tried, also with a changed property). Create the tables first, see 3c |
| `New-Object -ComObject Access.Application` scripts leave `MSACCESS.EXE -Embedding` behind | a script that ends on an error before `Quit`. Look with `Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'"` after every series of tests and stop your own `-Embedding` ones |
| `Build-FromSource.ps1` prints *Done* in the build log but the script itself does not return (seen once, for 30 minutes, after a 3-minute build) | the hidden Access did not quit. Look at the `-Embedding` processes and stop them (`Stop-Process`); the script then carries on. The build log and the database are complete at that point |

## 5. Compare a live database with the source

Opening the live file runs its startup form and AutoExec (a connection error window, a login form). On a **copy**:

```powershell
.\tools\vcs\Export-DatabaseText.ps1 -Database C:\temp\copy.mdb -Out C:\temp\live -What code -AllowBypass
.\tools\vcs\Export-DatabaseText.ps1 -Database C:\temp\copy.mdb -Out C:\temp\live -What forms -Skip BadForm
.\tools\vcs\Export-DatabaseText.ps1 -Database C:\temp\copy.mdb -Out C:\temp\live -What reports
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

## 9. Finding why Access refuses an object

When an object imports but will not save, or a build stops on an error that names nothing useful, do not guess: shrink it.
`tools/vcs/Shrink-FormText.ps1 -FormFile <form text> -Work <folder>` removes lines until the smallest text is left that still
gets refused. Run it again on the corrected text for a second cause. What worked, in the order we did it:

1. reproduce in a scratch database of the same format (`DBEngine.CreateDatabase(path, locale, 64)` is format 10, 128 is `.accdb`);
   a control test with a form that is known to be fine proves the test setup itself
2. try the obvious suspects one at a time (properties, controls) - cheap, but it rules out single causes only
3. let the shrink tool find the combination; then check the finding with four targeted variants (with/without each part)
4. only then change the client source, and rebuild the whole project once to confirm

Compile errors after a build from source are usually real findings, not build problems (for example a dead event handler that
referred to a field that no longer exists; in the original database it compiled because the form kept old type information).
