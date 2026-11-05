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

   It reads the stamps and hashes of [libstamp.py](../scripts/libstamp.py) (README, *Module stamps*), so formatting and the case
   of identifiers do not count as a difference. One section per situation, nothing is changed:

   | Section | Meaning | What you do |
   |---|---|---|
   | IDENTICAL | same as the library | nothing |
   | CLIENT BEHIND | untouched client copy, the library has moved on | update it (`copy`) |
   | CHANGED LOCALLY | the client copy no longer matches its own stamp | read the change: take it into the library, or discard it |
   | NO STAMP, DIFFERS | an old or foreign copy without a stamp | judge it like CHANGED LOCALLY; the counts show how far it is |
   | LOCAL ONLY | not in the library | put it in the library, or rename it with the client prefix |

3. Decide per module, then update the ones you chose:

   ```
   python -I scripts/clientsync.py copy <client source>/modules omDateFunctions omFileFunctions
   ```

   `copy` touches only the names you give. Class modules and standard modules are copied as they are (the same file format).
4. A module that the library now lacks something for, or needs a table, form or class the client does not have: follow the
   `uses` of the module (the compile in step 5 names what is missing; for example `omDateFunctions` needs `omTimeRange`).
5. Build and compile the client project from source, in a copy, as described in [BUILD-AND-TEST.md](BUILD-AND-TEST.md)
   (sections 2, 3 and 3c). A compile error after a refresh is a finding to look at, not noise.

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
