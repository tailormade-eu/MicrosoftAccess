<#
.SYNOPSIS
  Template: test VBA from this library in a hidden Access, through COM, one line of output per check.

.DESCRIPTION
  1. copy a built database (see tools/vcs/Build-FromSource.ps1 or the library build) to a work folder
  2. open it hidden, load a small helper module (tstHulp) that calls the code under test and returns TEXT
  3. $acc.Run('Function', args) and compare the text with what you expect: "OK name" / "FOUT name expected .. got .."
  4. quit Access; stop only the Access processes this script started

  Copy this file, replace the helper code and the Chk lines. A test that is green on its first run proves nothing until
  you have seen it go red (break the code on purpose, run, expect red, restore).

  Traps (all hit while building this library):
    - LoadFromText reads Windows-1252 and CRLF: write the helper with that encoding. A UTF-8 BOM shows up as code.
    - A .cls cannot be loaded with LoadFromText (class header). Standard modules only; the add-in build loads classes.
    - $acc.Run does not accept Module.Proc, and a function that returns an EMPTY string makes PowerShell fail with
      "Value cannot be null (Parameter 'ptr')": return "[empty]" or a count instead of "".
    - 'Dim x As New Class' creates the object again the moment you test 'x Is Nothing'. Use Set x = New ... in tests.
    - A modal window in a hidden Access (a MsgBox, an error, the VBE on a runtime error) blocks everything. The wait
      loop below reads the window text through UI Automation before it gives up.
    - A new report has to be saved (DoCmd.Save) before it is closed, or Access asks for a name and waits.
    - CurrentDb() is limited; reuse one object in loops.
#>
param([Parameter(Mandatory)] [string] $BuiltDatabase, [string] $WorkFolder = (Join-Path $env:TEMP 'access-test'))

if (Test-Path $WorkFolder) { Remove-Item $WorkFolder -Recurse -Force }
New-Item -ItemType Directory $WorkFolder | Out-Null
Copy-Item $BuiltDatabase (Join-Path $WorkFolder 't.accdb') -Force
$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue).Id

$job = Start-Job -ArgumentList $WorkFolder -ScriptBlock {
    param($work)
    $out = @(); $acc = $null
    function Chk($name, $got, $expected) {
        $ok = (("$got").Trim() -eq "$expected")
        $script:out += ('{0,-4} {1}' -f $(if ($ok) { 'OK' } else { 'FOUT' }), $name) + $(if (-not $ok) { "  expected [$expected] got [$got]" } else { '' })
    }
    function Rn($f, $a = @()) {
        try { switch ($a.Count) { 0 { '' + $acc.Run($f) } 1 { '' + $acc.Run($f, $a[0]) } default { '' + $acc.Run($f, $a[0], $a[1]) } } }
        catch { if ($_.Exception.Message -like '*ptr*') { '[empty]' } else { 'ERROR: ' + $_.Exception.Message } }
    }
    try {
        $acc = New-Object -ComObject Access.Application
        $acc.Visible = $false; $acc.AutomationSecurity = 1
        $acc.OpenCurrentDatabase((Join-Path $work 't.accdb'))
        $helper = @"
Attribute VB_Name = "tstHulp"
Option Explicit
Public Function TstExample(x As String) As String
    TstExample = UCase(x)      ' call the library code under test here
End Function
"@
        $file = Join-Path $work 'tstHulp.txt'
        [IO.File]::WriteAllText($file, ($helper -replace "`r?`n", "`r`n"), [Text.Encoding]::GetEncoding(1252))
        $acc.LoadFromText(5, 'tstHulp', $file)
        Chk 'example' (Rn 'TstExample' @('abc')) 'ABC'
    } catch { $out += 'ERROR: ' + $_.Exception.Message }
    finally { if ($acc) { try { $acc.CloseCurrentDatabase() } catch {}; try { $acc.Quit(2) } catch {} } }
    $out
}
if (Wait-Job $job -Timeout 120) { Receive-Job $job } else {
    'TIMEOUT: windows of the test Access:'
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $mine = @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'" | Where-Object { $_.CommandLine -match 'Embedding' } | ForEach-Object { $_.ProcessId })
    foreach ($k in [Windows.Automation.AutomationElement]::RootElement.FindAll('Children', [Windows.Automation.Condition]::TrueCondition)) {
        if ($mine -contains $k.Current.ProcessId -and $k.Current.ClassName -ne 'OMain') {
            "  [$($k.Current.ClassName)] $($k.Current.Name)"
            foreach ($c in $k.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)) { if ($c.Current.Name -and $k.Current.ClassName -eq '#32770') { '     ' + ($c.Current.Name -replace "`r?`n", ' / ') } }
        }
    }
    Stop-Job $job
}
Start-Sleep 2
# stop only what this script started: new since the start AND started through COM (-Embedding). Never another instance of the user.
foreach ($c in @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'")) { if (($before -notcontains $c.ProcessId) -and ($c.CommandLine -match '-Embedding')) { Stop-Process -Id $c.ProcessId -Force -ErrorAction SilentlyContinue } }
