<#
.SYNOPSIS
  Exports the objects of a COPY of a database to text files (SaveAsText), to compare a live database with source.

.DESCRIPTION
  The startup form and the AutoExec macro are skipped by holding the Shift key while the database is opened
  (the bypass key must be allowed in the copy: see -AllowBypass). Use it on a COPY, never on the live file.

  -What code     modules, queries, macros and a list of the tables
  -What forms    forms (one Access session; export in a separate run from the reports)
  -What reports  reports
  -Skip          object names to leave out. A form with a control that crashes Access when it is touched through
                 COM (RPC failed 0x800706BE) has to be skipped; the run continues without it.

  The text is Windows-1252 as SaveAsText writes it. Compare code with the header (VERSION / Attribute lines) and
  case differences ignored; the library's scripts/libstamp.py normalises text in the same way.

.EXAMPLE
  .\Export-DatabaseText.ps1 -Db C:\temp\copy.mdb -Out C:\temp\export -AllowBypass
  .\Export-DatabaseText.ps1 -Db C:\temp\copy.mdb -Out C:\temp\export -What forms -Skip Flow
#>
param(
    [Parameter(Mandatory)] [string] $Db,
    [Parameter(Mandatory)] [string] $Out,
    [ValidateSet('code', 'forms', 'reports')] [string] $What = 'code',
    [string[]] $Skip = @(),
    [switch] $AllowBypass
)
if ($AllowBypass) {
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $d = $dbe.OpenDatabase($Db)
    try { $d.Properties.Item('AllowBypassKey').Value = $true } catch { $d.Properties.Append($d.CreateProperty('AllowBypassKey', 1, $true)) }
    $d.Close()
}
Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class KB { [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, int e); }'
New-Item -ItemType Directory $Out -Force | Out-Null
$log = Join-Path $Out "_log_$What.txt"; Set-Content $log "start $(Get-Date -Format s)"
$acc = $null
try {
    [KB]::keybd_event(0x10, 0, 0, 0)                      # Shift down
    $acc = New-Object -ComObject Access.Application
    $acc.Visible = $false; $acc.AutomationSecurity = 1
    try { $acc.OpenCurrentDatabase($Db) } finally { [KB]::keybd_event(0x10, 0, 2, 0) }   # Shift up
    function Export-Set($type, $folder, $names) {
        $dir = Join-Path $Out $folder; New-Item -ItemType Directory $dir -Force | Out-Null
        $n = 0
        foreach ($name in $names) {
            if ($Skip -contains $name) { Add-Content $log "skipped: $name"; continue }
            try { $acc.SaveAsText($type, $name, (Join-Path $dir (($name -replace '[\\/:*?"<>|]', '_') + '.txt'))); $n++ } catch { Add-Content $log "ERROR $folder/$name : $($_.Exception.Message)" }
        }
        Add-Content $log "$folder : $n of $(@($names).Count)"
    }
    switch ($What) {
        'code' {
            Export-Set 5 'modules' @($acc.CurrentProject.AllModules | ForEach-Object { $_.Name })
            Export-Set 1 'queries' @($acc.CurrentDb().QueryDefs | ForEach-Object { $_.Name } | Where-Object { $_ -notlike '~*' })
            Export-Set 4 'macros' @($acc.CurrentProject.AllMacros | ForEach-Object { $_.Name })
            ($acc.CurrentDb().TableDefs | Where-Object { $_.Name -notlike 'MSys*' -and $_.Name -notlike '~*' } | ForEach-Object { $_.Name }) | Set-Content (Join-Path $Out '_tables.txt')
        }
        'forms'   { Export-Set 2 'forms'   @($acc.CurrentProject.AllForms   | ForEach-Object { $_.Name }) }
        'reports' { Export-Set 3 'reports' @($acc.CurrentProject.AllReports | ForEach-Object { $_.Name }) }
    }
    Add-Content $log "done $(Get-Date -Format s)"
} catch { Add-Content $log ('ERROR: ' + $_.Exception.Message) }
finally { if ($acc) { try { $acc.CloseCurrentDatabase() } catch {}; try { $acc.Quit(2) } catch {} } }
