<#
.SYNOPSIS
  Puts the rows of local tables from a COPY of the production file into a database that was built from source.

.DESCRIPTION
  A build from source restores every object, but a local table only gets its rows back when its data is in the export. Tables
  that are kept out of git on purpose (a table with a login, see docs/CLIENT-SYNC.md) and working data that is not exported
  stay empty. This script fills them from the production file, as the last step of a build:

      Build-FromSource.ps1 ...                                   -> the database built from source
      Restore-LocalTables.ps1 -BuiltDatabase <built> -ProductionFile <production file> -Tables SSMAA_ODBC_Tables,...

  - The production file is copied first and the copy is opened read-only: the production file itself is never opened.
  - Only the columns that exist in BOTH tables are copied (the built table may have new columns, the production table old ones);
    the others are reported by name. No value is ever printed.
  - Existing rows in the built table are deleted first (the built table is normally empty); -KeepExisting appends instead.
  - Writes into the built database only. Never point -BuiltDatabase at the production file.
  - -TablesFile takes one table name per line (# = comment), so a client repo can keep the list next to its source.

.EXAMPLE
  .\Restore-LocalTables.ps1 -BuiltDatabase C:\temp\build\App.mdb -ProductionFile D:\App.mdb -Tables SSMAA_ODBC_Tables,SSMAA_ODBC_Tables_E
#>
param(
    [Parameter(Mandatory)] [string] $BuiltDatabase,
    [Parameter(Mandatory)] [string] $ProductionFile,
    [string[]] $Tables = @(),
    [string] $TablesFile = '',
    [switch] $KeepExisting
)
if ((Resolve-Path $BuiltDatabase).Path -eq (Resolve-Path $ProductionFile).Path) { throw 'BuiltDatabase and ProductionFile are the same file: refused.' }
if ($TablesFile) { $Tables += @(Get-Content $TablesFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }) }
if (-not $Tables) { throw 'No tables: give -Tables or -TablesFile.' }

$copy = Join-Path $env:TEMP ('restore-source-' + [IO.Path]::GetFileName($ProductionFile))
Copy-Item $ProductionFile $copy -Force
$dbe = New-Object -ComObject DAO.DBEngine.120
$src = $dbe.OpenDatabase($copy, $false, $true)
$dst = $dbe.OpenDatabase($BuiltDatabase)
$bad = 0
try {
    foreach ($name in $Tables) {
      try {
        $st = try { $src.TableDefs.Item($name) } catch { $null }
        $dt = try { $dst.TableDefs.Item($name) } catch { $null }
        if (-not $st) { "{0}: NOT in the production file" -f $name; $bad++; continue }
        if (-not $dt) { "{0}: NOT in the built database (build it first; the table structure comes from the source)" -f $name; $bad++; continue }
        $srcCols = @($st.Fields | ForEach-Object { $_.Name })
        $dstCols = @($dt.Fields | ForEach-Object { $_.Name })
        $common = @($dstCols | Where-Object { $srcCols -contains $_ })
        $onlySrc = @($srcCols | Where-Object { $dstCols -notcontains $_ })
        $onlyDst = @($dstCols | Where-Object { $srcCols -notcontains $_ })
        $before = $dst.OpenRecordset("SELECT COUNT(*) FROM [$name]").Fields(0).Value
        if (-not $KeepExisting -and $before -gt 0) { $dst.Execute("DELETE FROM [$name]", 128) }   # 128 = dbFailOnError
        $list = ($common | ForEach-Object { "[$_]" }) -join ','
        # The IN clause fails through DAO from PowerShell ("Operation is not supported for this type of object", measured with
        # ACE 16), a temporary linked table works. INSERT ... SELECT also accepts explicit values for an AutoNumber column.
        $tmp = 'zzRestore_' + $name
        try { $dst.TableDefs.Delete($tmp) } catch {}
        $td = $dst.CreateTableDef($tmp); $td.Connect = ";DATABASE=$copy"; $td.SourceTableName = $name; $dst.TableDefs.Append($td)
        try { $dst.Execute("INSERT INTO [$name] ($list) SELECT $list FROM [$tmp]", 128) } finally { try { $dst.TableDefs.Delete($tmp) } catch {} }
        $nSrc = $src.OpenRecordset("SELECT COUNT(*) FROM [$name]").Fields(0).Value
        $nDst = $dst.OpenRecordset("SELECT COUNT(*) FROM [$name]").Fields(0).Value
        $ok = if ($KeepExisting) { $nDst -ge $nSrc } else { $nDst -eq $nSrc }
        $line = "{0}: production {1} rows -> built {2} rows ({3} columns copied)" -f $name, $nSrc, $nDst, $common.Count
        if ($onlySrc) { $line += "; only in production (not copied): " + ($onlySrc -join ',') }
        if ($onlyDst) { $line += "; only in built (left empty): " + ($onlyDst -join ',') }
        if (-not $ok) { $line += '  <-- ROW COUNT DIFFERS'; $bad++ }
        $line
      } catch { "{0}: FAILED ({1})" -f $name, $_.Exception.Message; $bad++ }
    }
} finally { $src.Close(); $dst.Close(); Remove-Item $copy -Force -ErrorAction SilentlyContinue }
if ($bad) { "problems: $bad"; exit 1 } else { 'all tables restored' }
