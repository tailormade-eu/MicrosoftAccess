<#
.SYNOPSIS
  Builds a database from a VCS export folder (the MSAccessVCS add-in) in a throw-away copy, in a hidden Access.

.DESCRIPTION
  Copies the source folder to a work folder, optionally strips what stops an unattended build, creates an empty
  database and calls the add-in's Build. The source folder and the real database are never touched.

  What the script does to the COPY (all switches are for test builds only):
    - "VCS Build Path" in proj-properties.json is pointed at the work folder. Without that, the add-in asks
      "Build Name Conflict ... build to the source defined filename?" and OK would build over the REAL database.
      If that window still appears the script presses Cancel, never OK.
    - BreakOnError is turned off in vcs-options.json, so an error goes to the build log instead of stopping the
      VBE on a Stop that nobody sees in a hidden Access.
    - Linked (ODBC) tables, db-connection.json and pass-through queries with an ODBC connection are removed
      (-DropLinkedTables): without a connection Access asks for a data source for each of them.
    - -DropThemes removes the themes folder; -DropObjects removes single objects by name.

  Done = the add-in log is stable, or Access used no CPU for 45 s and the database is larger than 1 MB.
  Only Access processes started by this script (command line contains -Embedding) are stopped at the end.

.EXAMPLE
  .\Build-FromSource.ps1 -SourceFolder D:\repos\MyApp\MyApp.accdb.src -DatabaseName MyApp.accdb -DropLinkedTables
#>
param(
    [Parameter(Mandatory)] [string] $SourceFolder,
    [Parameter(Mandatory)] [string] $DatabaseName,
    [string] $WorkFolder = (Join-Path $env:TEMP ('vcs-build\' + [IO.Path]::GetFileNameWithoutExtension($DatabaseName))),
    [int] $TimeoutSec = 900,
    [switch] $DropLinkedTables,
    [switch] $DropThemes,
    [string[]] $DropObjects = @(),   # e.g. reports/SomeReport (all files with that name)
    [switch] $KeepBreakOnError
)

# the installed add-in: %AppData%\MSAccessVCS, or, when AppData is redirected to OneDrive, under <OneDrive folder>\AppData\Roaming
$addin = $null
$roots = @([Environment]::GetFolderPath('ApplicationData'))
foreach ($od in @($env:USERPROFILE, $env:OneDrive, $env:OneDriveCommercial)) {   # a variable can be empty on a pc: skip it
    if ($od) { $roots += (Join-Path $od 'AppData\Roaming'); $roots += (Join-Path $od 'OneDrive\AppData\Roaming') }
}
foreach ($root in $roots) {
    if (Test-Path (Join-Path $root 'MSAccessVCS\Version Control.accda')) { $addin = Join-Path $root 'MSAccessVCS\Version Control'; break }
}
if (-not $addin) { throw 'The VCS add-in (MSAccessVCS\Version Control.accda) was not found under AppData; install it or pass the folder.' }
if (Test-Path $WorkFolder) { Remove-Item $WorkFolder -Recurse -Force }
New-Item -ItemType Directory $WorkFolder | Out-Null
$src = Join-Path $WorkFolder ($DatabaseName + '.src')
$db = Join-Path $WorkFolder $DatabaseName
robocopy $SourceFolder $src /E /XD .git logs /NFL /NDL /NJH /NJS | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy failed ($LASTEXITCODE)" }   # 0-7 are success codes

function Write-Utf8Bom($path, $text) { [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($true))) }

# 1. never build over the real database
$pp = Join-Path $src 'proj-properties.json'
if (Test-Path $pp) {
    $t = [IO.File]::ReadAllText($pp, [Text.Encoding]::UTF8)
    $t = [regex]::Replace($t, '("VCS Build Path":\s*")[^"]*(")', { param($m) $m.Groups[1].Value + $WorkFolder.Replace('\', '\\') + $m.Groups[2].Value })
    Write-Utf8Bom $pp $t
}
# 2. log errors instead of stopping the VBE
$vo = Join-Path $src 'vcs-options.json'
if ((-not $KeepBreakOnError) -and (Test-Path $vo)) {
    $t = [IO.File]::ReadAllText($vo, [Text.Encoding]::UTF8)
    $t2 = [regex]::Replace($t, '("BreakOnError":\s*)true', '${1}false')
    if ($t2 -ne $t) { Write-Utf8Bom $vo $t2; 'BreakOnError turned off in the copy' }
}
# 3. no connections in a test build
if ($DropLinkedTables) {
    $n = 0
    $tbl = Join-Path $src 'tbldefs'
    if (Test-Path $tbl) { foreach ($f in Get-ChildItem $tbl -Filter *.json) { if ((Get-Content $f.FullName -Raw) -match '"Description":\s*"Linked Table"') { Remove-Item $f.FullName -Force; $n++ } } }
    $dc = Join-Path $src 'db-connection.json'
    if (Test-Path $dc) { Remove-Item $dc -Force }
    $qd = Join-Path $src 'queries'
    $q = 0
    if (Test-Path $qd) {
        foreach ($f in Get-ChildItem $qd -Filter *.json) {
            if ((Get-Content $f.FullName -Raw) -match 'env:conn|ODBC;') {
                Remove-Item $f.FullName -Force
                $sql = [IO.Path]::ChangeExtension($f.FullName, '.sql'); if (Test-Path $sql) { Remove-Item $sql -Force }
                $q++
            }
        }
    }
    "removed from the copy: $n linked tables, $q pass-through queries, db-connection.json"
}
if ($DropThemes -and (Test-Path (Join-Path $src 'themes'))) { Remove-Item (Join-Path $src 'themes') -Recurse -Force; 'themes removed from the copy' }
foreach ($o in $DropObjects) {
    $p = Join-Path $src $o.Replace('/', '\')
    $dir = Split-Path $p; $leaf = Split-Path $p -Leaf
    if (Test-Path $dir) { foreach ($x in Get-ChildItem $dir -File | Where-Object { $_.BaseName -eq $leaf }) { Remove-Item $x.FullName -Force; "removed: $($x.Name)" } }
}

$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue).Id
$job = Start-Job -ArgumentList $db, $src, $addin, $TimeoutSec -ScriptBlock {
    param($db, $src, $addin, $TimeoutSec)
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $out = @(); $acc = $null
    function Mine { Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'" | Where-Object { $_.CommandLine -match 'Embedding' } | Sort-Object CreationDate -Descending | Select-Object -First 1 }
    try {
        $acc = New-Object -ComObject Access.Application
        $acc.Visible = $false; $acc.AutomationSecurity = 1
        $acc.NewCurrentDatabase($db)
        $null = $acc.Run("$addin.API", 'Build', $src)
        $logs = Join-Path $src 'logs'
        $start = Get-Date; $last = $null; $stable = 0; $cpuLast = -1; $cpuStill = 0
        while (((Get-Date) - $start).TotalSeconds -lt $TimeoutSec) {
            Start-Sleep -Seconds 3
            $mine = Mine
            if ($mine) {
                foreach ($k in [Windows.Automation.AutomationElement]::RootElement.FindAll('Children', [Windows.Automation.Condition]::TrueCondition)) {
                    if ($k.Current.ProcessId -eq $mine.ProcessId -and $k.Current.ClassName -eq '#32770') {
                        $texts = @(); $cancel = $null
                        foreach ($c in $k.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)) {
                            if ($c.Current.Name) { $texts += ($c.Current.Name -replace "`r?`n", ' / ') }
                            if ($c.Current.ControlType.ProgrammaticName -match 'Button' -and $c.Current.Name -eq 'Cancel') { $cancel = $c }
                        }
                        $out += "DIALOG '$($k.Current.Name)': $($texts -join ' | ')"
                        if ($cancel) { $out += '  -> Cancel'; $cancel.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() }
                    }
                }
            }
            $l = Get-ChildItem $logs -Filter 'Build_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            $sig = if ($l) { "$($l.Name)|$($l.Length)" } else { '' }
            if ($sig -ne '' -and $sig -eq $last) { $stable++ } else { $stable = 0 }
            $last = $sig
            if ([IO.File]::Exists($db) -and $stable -ge 8) { break }
            $proc = if ($mine) { Get-Process -Id $mine.ProcessId -ErrorAction SilentlyContinue }
            if ($proc) {
                $cpu = [math]::Round($proc.CPU, 1)
                if ($cpu -eq $cpuLast) { $cpuStill++ } else { $cpuStill = 0 }
                $cpuLast = $cpu
                if ($cpuStill -ge 15 -and [IO.File]::Exists($db) -and (Get-Item $db).Length -gt 1MB) { $out += '(done: Access idle for 45 s)'; break }
            }
        }
        $l = Get-ChildItem $logs -Filter 'Build_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($l) { $out += "--- build log: $($l.FullName)"; $out += (Select-String -Path $l.FullName -Pattern 'ERROR|Error \d+|CRITICAL|warning\(s\)|Done\.' | ForEach-Object { $_.Line }) }
    } catch { $out += 'ERROR: ' + $_.Exception.Message }
    finally { if ($acc) { try { $acc.CloseCurrentDatabase() } catch {}; try { $acc.Quit(2) } catch {} } }
    $out
}
if (Wait-Job $job -Timeout ($TimeoutSec + 60)) { Receive-Job $job } else { 'TIMEOUT'; Stop-Job $job }
Start-Sleep 3
foreach ($p in @(Get-Process MSACCESS -ErrorAction SilentlyContinue)) { if ($before -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } }
"database: $db (exists: $(Test-Path $db))"
$global:LASTEXITCODE = 0   # robocopy leaves 1 behind (= files copied), which is not a failure
