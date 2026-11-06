# Working on a client's Access project: keep the library and the client in step

The goal is **as little drift as possible** between this library and the copies in client databases. Every `om*` module
in a client project should be either identical to the library, or a known and recorded difference.

```
 library (this repo)  <----->  client source repo  ----->  client database
        ^                                                       |
        +------- improvements made while working for the client +
```

Principles (decisions of the owner):
- **`om*` = library.** A client project holds identical copies. Code that belongs to one client does not start with `om`;
  it gets a client prefix (for example `hbPrinterFunctions`), and is never deleted when a library module replaces it.
- **A library function that needs a table or form owns it** (`omPrinters`, `omFiles`, `omHelpFile`, `Calendar`): the function
  creates it. The client does not keep its own copy.
- **Improvements travel both ways.** If you fix or extend an `om*` module for a client, the fix goes into the library first
  (or at the same time), and the client gets the library version. A client copy that differs from the library is drift.
- **Nothing is changed in a client's real database** by this process. Work in the source repo and in copies (see
  [BUILD-AND-TEST.md](BUILD-AND-TEST.md)).

## At the start of work for a client

1. Update the library repo (`git pull`) and the client source repo. Work on a branch of the client repo.
2. Run the drift report on the client's modules folder:

   ```
   python -I scripts/clientsync.py report <client source>/modules
   ```

   It uses the stamps and hashes of [libstamp.py](../scripts/libstamp.py) (README, *Module stamps*) and, on top of that, a
   looser comparison of the code itself: outside string literals the case is ignored, runs of spaces and tabs count as one,
   indentation and blank lines are ignored. Inside a string literal nothing is touched, so `"A"` and `"a"` still differ. The
   stamp hash itself stays strict about spaces and blank lines, because `omLibraryStamp.bas` must calculate the same hash inside
   Access; that is why formatting-only copies are reported separately and not as a difference. One section per situation,
   nothing is changed:

   | Section | Meaning | What you do |
   |---|---|---|
   | IDENTICAL | same as the library | nothing |
   | IDENTICAL, NO STAMP | same code, but the client copy has no stamp line | `copy` it: with the stamp a later library change shows as CLIENT BEHIND, without it only as NO STAMP, DIFFERS |
   | FORMATTING ONLY | same code, other case of identifiers, other spaces or indentation, other blank lines | nothing, or `copy` to make the hash match again |
   | CLIENT BEHIND | untouched client copy, the library has moved on | update it (`copy`) |
   | CHANGED LOCALLY | the client copy no longer matches its own stamp | read the change. **A client may have changed a parameter in the om* module (a Const, a True/False, a server name, a default): that is a setting of the client and moves to its `ClientSettings_LoadOverride`, see "Settings of a client" below, before you `copy`.** Other changes: take them into the library, or discard them |
   | NO STAMP, DIFFERS | an old or foreign copy without a stamp | judge it like CHANGED LOCALLY; the counts show how far it is |
   | LOCAL ONLY | not in the library | put it in the library, or rename it with the client prefix |
   | CLASS ATTRIBUTES DIFFER | the class attributes (`VB_PredeclaredId` and the like) differ while the code is the same | `copy`, then check how the class is used |

   The last line of the report counts the library modules the client does not have; that is not an issue (a client takes what it uses).

3. Decide per module, then update the ones you chose:

   ```
   python -I scripts/clientsync.py copy <client source>/modules omDateFunctions omFileFunctions
   ```

   `copy` touches only the names you give, and only modules the client already has (for another name it says so and exits with 1; `add` puts a new module in and refuses one that exists). It warns when it overwrites a client copy that was CHANGED LOCALLY or NO STAMP, DIFFERS: take such a change into the library first. A name that is not in the library gives "not in the library" and exit code 1.

   To keep a client on the whole library use `python -I scripts/clientsync.py all <client source>/modules`: it adds what is missing and refreshes what differs, skips and lists every client copy that holds an unjudged change (exit code 1), and leaves LOCAL ONLY modules alone. Run `report` first; run steps 5 and 6 after it. Class modules and standard modules are copied as they are (the same file format).
4. A module that the library now lacks something for, or needs a table, form or class the client does not have: follow the
   `uses` of the module (the compile in step 5 names what is missing; for example `omDateFunctions` needs `omTimeRange`).
5. Build and compile the client project from source, in a copy, as described in [BUILD-AND-TEST.md](BUILD-AND-TEST.md)
   (sections 2, 3 and 3c). A compile error after a refresh is a finding to look at, not noise.
6. **Check what the compiler cannot see.** A green compile does not look inside field expressions, queries and macros:

   ```
   python -I scripts/expressionrefs.py report <client source>
   ```

   It lists the functions that forms, reports, queries (also a SQL `RecordSource`) and macros call by name (and the ones in a conditional format), and that no module defines (`MISSING`), or
   defines in a way an expression cannot call (`PRIVATE`, `SUB`, `CLASS`). Run it on the project before the refresh and after:
   anything new in the second list is something the refresh removed. Findings that were already in the first list are the
   client's own and are not caused by the refresh (still worth telling the owner). Case in point: a refresh replaced a module
   that held `RoundHalfUp`; two invoice reports called it in their field expressions; the compile stayed green, this check
   reports it (`Roundhalfup  - 8 uses in 2 objects`). The built-in function list in the script is not complete: a real Access
   function that shows as MISSING gets added to `BUILTINS`. Tested in `scripts/test_expressionrefs.py`.

## Measured: the whole library in one client

On 2026-11-06 all 43 library modules that a client project did not have were copied into a copy of its source (131 module
files), built from source as `.mdb` and compiled. The build had no errors; the compile stopped once, in `omTesseractOCR`
(`IWshRuntimeLibrary.WshShell`: the *Windows Script Host Object Model* reference was missing). With that one reference added
(`Compile-InVbe.ps1 -AddReference '{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}'`) the whole project compiled. So the whole library
fits in that client; a client takes the references of `src/VCS/MicrosoftAccess.accdb.src/vbe-references.json` with it.
Compiling is not running: the modules were not run in that copy, and the built file is 37 MB (with the stand-in tables; a build without the 43 modules was not measured).

## Settings of a client: never in an om* module

A value that one client sets and another does not (which login, development mode, a language, the address of a Seq server)
does not belong in a library module that is copied: the next copy overwrites it, and the client then runs with another value
(this happened once: a refresh of `omUserFunctions` replaced the Windows login of a client by the login form). The way it works:

- The library holds the standard value of every setting in `omGlobalDeclarations.ClientSettings_Load`.
- A client changes a value in its own `Public Sub ClientSettings_LoadOverride()` (a standard module of the client, no
  arguments, plain assignments only). The library finds it by name with `Application.Run`; a client without the
  procedure just gets the standard values (error 2517 is the only error that is ignored).
- Every module that reads a setting calls `omGlobalDeclarations.EnsureClientSettings` first. After a reset of the VBA project
  (an unhandled error, `End`) the variables are empty again; the next call loads them again, from the same two places.
- The login has no standard: `UseWindowsAuthentication` and `UseLoginForm` are both False until the client sets exactly one,
  and `AuthenticateUser` says so instead of choosing a login screen.
- An error that is raised inside the override is not given back to the library: Access shows its own run-time error
  dialog (measured). Keep the override to plain assignments.

Example (a client with a Windows login, in the client's own `StartupFunctions.bas`):

```vba
Public Sub ClientSettings_LoadOverride()
    omUserFunctions.UseWindowsAuthentication = True
    omUserFunctions.UseLoginForm = False
End Sub
```

Tested in `tstClientSettings` (in Access) and `scripts/test_library_rules.py` (no True/False `Const` in an om* module).
A client may have overwritten a parameter in its own copy of an om* module (the old `Const UseWindowsAuthentication` / `UseLoginForm`
in `omUserFunctions` were exactly that). The report then says CHANGED LOCALLY. Before every `copy` or `all` look at what the client
changed (`git diff` of the client copy against the library file): a value that is the client's own goes into its
`ClientSettings_LoadOverride`, and only then is the module replaced. Without this step the refresh silently resets the value
(that is the regression that led to this section: a client got another login screen after a refresh).

## Local tables with configuration or data: always export them with their data

A database built from source gets its objects back, but a local table only gets its rows back when the export options export
the data of that table (`TablesToExportData` in `vcs-options.json`). A local table that holds configuration or data of its
own (translations, field lists, defaults, the connection table of the linking code) must therefore always be exported with
its data, or a database built from the source starts with an empty table and the app does not work.

- Check it for a client: `tools/vcs/List-LocalTables.ps1 -Database <copy of the production file> -SourceFolder <client source>`.
  Per local table: rows, whether its data is exported, whether a column looks like a password. `LOOK` = has rows and is not
  exported: decide per table (configuration or data: add it to the export; a working copy such as the `T_` tables of
  `omTableConnector`: leave it out).
- **Exception: a table with a login or other secret** (a column such as `SQLPassword`) does NOT go into git, not even in a
  private repo (credentials policy). Put the table in `.gitignore` (`tables/<name>*.xml`), keep its structure in the repo, and
  keep its data in the production file or in a secret store. A database built from the source then has this table empty: fill
  it from the production file (or from the secret store) before the app is used.
- **A build from source then needs one more step: put the table data back.** `tools/vcs/Restore-LocalTables.ps1 -BuiltDatabase
  <built> -ProductionFile <production file> -Tables <names>` (or `-TablesFile`, one name per line, so a client repo can keep the
  list next to its source) copies the rows of those tables from a COPY of the production file into the built database: the
  tables that are kept out of git (the login table) and the working data. Only the columns that exist in both tables are copied
  (the others are listed by name), no value is printed, it stops with an error when a table is missing or the row counts differ.
  Measured on a client: 4 tables, 1.237 rows, the row counts matched and so did the sum of the value lengths of the copied columns (a weak check: it does not compare the values one by one).
  Note: a table whose structure is newer in the source than in production (new columns) gets those columns empty; check that the
  code that reads them can handle that before the built database is used.
- A build made from a working folder that still holds an ignored data file (`tables/<name>.xml` on disk, not in git) is not a
  test of a clean clone: it restores that table from the file. Test from a fresh clone, or delete the ignored files first.
- Never replace the production file by a database built from source. Merge the changed objects into a copy of the production
  file, or use the built database (after `Restore-LocalTables.ps1`) as a test file.

## During the work

- Edit an `om*` module in the library repo, not in the client copy. If you did edit it in the client (quicker, or you were
  in the database), the next report says CHANGED LOCALLY: move the change over before you forget it.
- New reusable code gets the `om` prefix and goes into the library, with a test and the stamp
  (`python -I scripts/libstamp.py apply`).

## At the end of work

1. Run the report again. The aim is: no CHANGED LOCALLY and no NO STAMP, DIFFERS, and every LOCAL ONLY decided.
2. Library: `python -I scripts/libstamp.py apply`, `check`, commit and (when the owner says so) push.
3. Client repo: commit on the branch. Record in the client's notes which library state it now has (date of the newest stamp in
   `library-stamps.txt`) and which differences are intentional.
4. The published client database is a separate step with its own approval. The report tells you what changed; the owner
   decides when the client gets a new build.

## What this process does not do

- It does not run on its own. A schedule or a git hook would be new configuration and needs the owner's approval first
  (what it does, why, default, who reads it, when it goes away).
- It compares `om*` modules only. Forms, reports, queries and tables of the client are the client's own.
- A matching hash means the text is the same, not that the module works in that client: compile and, where it matters, run.

The comparison is tested in `scripts/test_clientsync.py` (`python -I scripts/test_clientsync.py`).
