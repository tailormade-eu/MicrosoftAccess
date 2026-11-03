<#
.SYNOPSIS
  Makes local stand-in tables for the linked tables of a client database, as tbldefs\<name>.xml files for a VCS source folder.

.DESCRIPTION
  A client project that links to a server (SQL Server) cannot be built and compiled on a pc without that server: Build-FromSource.ps1
  -DropLinkedTables removes the links, and a form bound to a removed table has no field list. Code such as Me.CreateDate then
  fails to compile - not because the code is wrong but because the form was read in before its table existed.

  The order matters: tables must exist before forms and reports are read in. The VCS add-in builds local tables from
  tbldefs\<name>.xml before it reads the forms, so this script writes such a file per linked table:
    1. opens a database that still has the links (a copy of the real database; Access keeps the field list of a link)
    2. creates a local table with the same field names and types in a scratch .accdb
    3. exports its structure with ExportXML into <SourceFolder>\tbldefs\<name>.xml

  Use it on a COPY of the source folder, then build that copy with Build-FromSource.ps1 -DropLinkedTables. The stand-in tables hold
  no data and no keys. The field list is the one Access cached when the link was last refreshed: a field that was added or
  renamed on the server since then is missing, and code that uses it fails to compile (a real finding, not a build problem).

.EXAMPLE
  .\New-StandInTables.ps1 -Template C:\temp\copy-of-real.mdb -SourceFolder C:\temp\src-copy
#>
param(
    [Parameter(Mandatory)] [string] $Template,
    [Parameter(Mandatory)] [string] $SourceFolder,
    [string] $WorkFolder = (Join-Path $env:TEMP 'standin-tables')
)
if (Test-Path $WorkFolder) { Remove-Item $WorkFolder -Recurse -Force }
New-Item -ItemType Directory $WorkFolder | Out-Null
$scratch = Join-Path $WorkFolder 'standin.accdb'
$tbldefs = Join-Path $SourceFolder 'tbldefs'
if (-not (Test-Path $tbldefs)) { throw "No tbldefs folder in $SourceFolder" }

# 1 + 2: local tables with the fields of the links
$e = New-Object -ComObject DAO.DBEngine.120
$dst = $e.CreateDatabase($scratch, ';LANGID=0x0409;CP=1252;COUNTRY=0', 128)
$src = $e.OpenDatabase($Template, $false, $true)
$names = @(); $failed = @()
foreach ($t in $src.TableDefs) {
    if (-not $t.SourceTableName) { continue }
    try {
        $td = $dst.CreateTableDef($t.Name)
        foreach ($f in $t.Fields) {
            $nf = if ($f.Type -eq 10 -or $f.Type -eq 12) { $td.CreateField($f.Name, $f.Type, [math]::Max(1, [math]::Min(255, $f.Size))) } else { $td.CreateField($f.Name, $f.Type) }
            $td.Fields.Append($nf)
        }
        $dst.TableDefs.Append($td); $names += $t.Name
    } catch { $failed += "$($t.Name): $($_.Exception.Message)" }
}
$src.Close(); $dst.Close()

# 3: structure as XML schema
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false; $acc.AutomationSecurity = 1
try {
    $acc.OpenCurrentDatabase($scratch)
    foreach ($n in $names) { $acc.ExportXML(0, $n, $null, (Join-Path $tbldefs ($n + '.xml'))) }
} finally { try { $acc.CloseCurrentDatabase() } catch {}; try { $acc.Quit(2) } catch {} }
"stand-in tables written to $tbldefs : $($names.Count), failed: $($failed.Count)"
$failed | Select-Object -First 10
