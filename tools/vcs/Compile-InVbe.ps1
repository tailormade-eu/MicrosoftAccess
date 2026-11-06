<#
.SYNOPSIS
  Compiles a database in a hidden Access through the VBE (Debug > Compile) and reports the module and line of the first error.

.DESCRIPTION
  Use it on a COPY. It sets AllowBypassKey in the copy and holds Shift while opening, so a startup form (a login, an
  AutoExec message box) does not block the run. The compile command opens a message box on an error: this script reads it
  through UI Automation and presses OK. Output:
    compile control enabled: True          the Compile menu item was available before the run
    ERRORPOS module=<name> line=<n> ...    the first error, with the lines around it
    no active code pane ... / enabled after: False
                                           no error: the Compile item is greyed out after a successful compile
  On a timeout it lists the windows of the test Access (a modal window is the usual cause).

  -AddReference <guid>[,<guid>...] adds those references to the copy before the compile (for example the Windows Script Host
  library {F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}), so you can see which modules need which reference without a new build.

.EXAMPLE
  .\Compile-InVbe.ps1 -Database C:\temp\copy.mdb
  .\Compile-InVbe.ps1 -Database C:\temp\copy.mdb -AddReference '{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}'
#>
param([Parameter(Mandatory)][string]$Database, [string[]]$AddReference = @())
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$dbe = New-Object -ComObject DAO.DBEngine.120
$d0 = $dbe.OpenDatabase($Database)
try { $d0.Properties.Item('AllowBypassKey').Value = $true } catch { $d0.Properties.Append($d0.CreateProperty('AllowBypassKey', 1, $true)) }
$d0.Close()
$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue).Id
$job = Start-Job -ArgumentList $Database, $AddReference -ScriptBlock {
    param($db, $refs)
    $out = @(); $acc = $null
    try {
        $acc = New-Object -ComObject Access.Application
        $acc.Visible = $false; $acc.AutomationSecurity = 1
        Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class KB { [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, int e); }'
        [KB]::keybd_event(0x10, 0, 0, 0)
        try { $acc.OpenCurrentDatabase($db) } finally { [KB]::keybd_event(0x10, 0, 2, 0) }
        foreach ($g in $refs) {
            try { $null = $acc.References.AddFromGuid($g, 0, 0); $out += "reference added: $g" } catch { $out += "reference NOT added: $g ($($_.Exception.Message))" }
        }
        $acc.VBE.MainWindow.Visible = $true
        $ctl = $acc.VBE.CommandBars.FindControl(1, 578)
        $out += 'compile control enabled: ' + $ctl.Enabled
        $ctl.Execute()
        $out += 'COMPILE-RETURNED'
        try {
            $pane = $acc.VBE.ActiveCodePane
            $sl = 0; $sc = 0; $el = 0; $ec = 0
            $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
            $m = $pane.CodeModule
            $out += "ERRORPOS module=$($m.Name) line=$sl col=$sc"
            $from = [math]::Max(1, $sl - 4)
            $out += $m.Lines($from, 9)
        } catch { $out += 'no active code pane: ' + $_.Exception.Message }
        $out += 'compile control enabled after: ' + $ctl.Enabled
    } catch { $out += 'ERROR: ' + $_.Exception.Message }
    finally { if ($acc) { try { $acc.CloseCurrentDatabase() } catch {}; try { $acc.Quit(2) } catch {} } }
    $out
}
$start = Get-Date
while ($job.State -eq 'Running' -and ((Get-Date) - $start).TotalSeconds -lt 75) {
    Start-Sleep -Seconds 2
    $mine = @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'" | Where-Object { $before -notcontains $_.ProcessId } | ForEach-Object { $_.ProcessId })
    foreach ($k in [Windows.Automation.AutomationElement]::RootElement.FindAll('Children', [Windows.Automation.Condition]::TrueCondition)) {
        if ($mine -contains $k.Current.ProcessId -and $k.Current.ClassName -eq '#32770') {
            $texts = @(); $ok = $null
            foreach ($c in $k.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)) {
                if ($c.Current.Name) { $texts += ($c.Current.Name -replace "`r?`n", ' / ') }
                if ($c.Current.ControlType.ProgrammaticName -match 'Button' -and $c.Current.Name -eq 'OK') { $ok = $c }
            }
            "DIALOG '$($k.Current.Name)': $($texts -join ' | ')"
            if ($ok) { $ok.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() }
        }
    }
}
if ($job.State -eq 'Running') {
    'TIMEOUT: windows of the test Access:'
    $mine = @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'" | Where-Object { $before -notcontains $_.ProcessId } | ForEach-Object { $_.ProcessId })
    foreach ($k in [Windows.Automation.AutomationElement]::RootElement.FindAll('Children', [Windows.Automation.Condition]::TrueCondition)) {
        if ($mine -contains $k.Current.ProcessId) {
            "  [$($k.Current.ClassName)] $($k.Current.Name)"
            if ($k.Current.ClassName -ne 'OMain') { foreach ($c in $k.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)) { if ($c.Current.Name) { '     ' + ($c.Current.Name -replace "`r?`n", ' / ') } } }
        }
    }
    Stop-Job $job } else { Receive-Job $job }
Start-Sleep 2
# stop only what this script started: new since the start AND started through COM (-Embedding). Never another instance of the user.
foreach ($c in @(Get-CimInstance Win32_Process -Filter "Name='MSACCESS.EXE'")) { if (($before -notcontains $c.ProcessId) -and ($c.CommandLine -match '-Embedding')) { Stop-Process -Id $c.ProcessId -Force -ErrorAction SilentlyContinue } }
