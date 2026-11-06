<#
.SYNOPSIS
  Lists the local tables of an Access database (a COPY of it, opened read-only) with their row counts, and says per table whether
  the VCS export options export its data.

.DESCRIPTION
  A build from source only restores the data of the tables that are exported with their data. Every local table that holds
  configuration or data of its own must be on that list; otherwise a database built from the source starts with an empty table.
  The script shows, per local table (not linked, not a system table):
    Rows      number of records
    Exported  yes when the table is in "TablesToExportData" of <SourceFolder>\vcs-options.json
    Secret?   a column name that looks like a password, token or secret: such a table must NOT go into git (credentials policy);
              keep the structure in the repo and the data in the production file (or in a secret store)
    Check     what to decide: LOOK  = has rows, not exported (config/data? then add it to the export; a working copy: ignore)
                              EMPTY = no rows, nothing to export
                              OK    = exported, or has no rows
                              SECRET= has a secret column: never export the data to git
  Tables with the prefix T_ or XD_ are normally working copies (omTableConnector); they are listed, not judged.

.EXAMPLE
  .\List-LocalTables.ps1 -Database C:\temp\copy.mdb -SourceFolder D:\repos\App\App.mdb
#>
param(
    [Parameter(Mandatory)] [string] $Database,
    [string] $SourceFolder = ''
)
$exported = @{}
if ($SourceFolder) {
    $vo = Join-Path $SourceFolder 'vcs-options.json'
    if (Test-Path $vo) {
        $j = [IO.File]::ReadAllText($vo, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF) | ConvertFrom-Json
        if ($j.Options.TablesToExportData) { $j.Options.TablesToExportData.PSObject.Properties | ForEach-Object { $exported[$_.Name] = $true } }
    }
}
$cp = Join-Path $env:TEMP ('list-local-tables-' + [IO.Path]::GetFileName($Database))
Copy-Item $Database $cp -Force
$dbe = New-Object -ComObject DAO.DBEngine.120
$db = $dbe.OpenDatabase($cp, $false, $true)
try {
    $rows = foreach ($t in $db.TableDefs) {
        if ($t.Name -like 'MSys*' -or $t.Name -like '~*' -or $t.Connect) { continue }
        $n = try { $db.OpenRecordset("SELECT COUNT(*) FROM [$($t.Name)]").Fields(0).Value } catch { -1 }
        $cols = ($t.Fields | ForEach-Object { $_.Name }) -join ','
        $secret = $cols -match '(?i)pass|pwd|secret|token'
        $isExported = $exported.ContainsKey($t.Name)
        $check = if ($secret -and $n -gt 0) { 'SECRET' } elseif ($n -le 0) { 'EMPTY' } elseif ($isExported) { 'OK' } else { 'LOOK' }
        [pscustomobject]@{ Table = $t.Name; Rows = $n; Exported = $(if ($isExported) { 'yes' } else { '-' }); 'Secret?' = $(if ($secret) { 'yes' } else { '' }); Check = $check }
    }
} finally { $db.Close() }
$rows | Where-Object { $_.Rows -gt 0 -or $_.Exported -eq 'yes' } | Sort-Object Check, @{ Expression = 'Rows'; Descending = $true } | Format-Table -AutoSize | Out-String -Width 160
"empty local tables not listed: $(@($rows | Where-Object { $_.Rows -le 0 -and $_.Exported -ne 'yes' }).Count)"
Remove-Item $cp -Force -ErrorAction SilentlyContinue
