<#
.SYNOPSIS
  Template: test the Excel modules of this library (MicrosoftExcel/src/modules) in a hidden Excel, through COM.

.DESCRIPTION
  Imports the modules into a new workbook, adds a helper module that calls them and returns TEXT, runs it with
  $xl.Run and compares. Needs, on this pc, in Excel: Trust Center > Macro Settings > "Trust access to the VBA project
  object model" (registry: HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1). This script sets it for the
  run and puts the old value back.

  Traps:
    - Import files as Windows-1252 with CRLF, no BOM (the library writes them that way).
    - A class module cannot be filled with AddFromString/AddFromFile when the text still has its header
      (VERSION 1.0 CLASS ... END and the Attribute lines): strip them first (see ReadSourceWithoutHeader in
      MicrosoftExcel/src/modules/omExcelVersionControl.bas).
    - Replacing a module that other code calls: remove-and-import leaves the new copy with another name (Module1)
      while the old one is still referenced. Replace the CODE of the existing module instead.
    - Workbooks.Open(file, Format:=xlCSV): Format is the DELIMITER argument, not the file type. Leave it out.
    - A function that returns "" fails through Run (Parameter 'ptr'); and RecordCount is -1 on a dynamic cursor: loop.
    - A compile error makes a VBE window appear and everything waits. The timeout branch reads its text.
#>
param(
    [string] $ModuleFolder = (Join-Path $PSScriptRoot '..\..\MicrosoftExcel\src\modules'),
    [string[]] $Modules = @('omExcelApplicationFunctions'),
    [string] $WorkFolder = (Join-Path $env:TEMP 'excel-test')
)
$key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$old = (Get-ItemProperty $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
if (-not (Test-Path $key)) { New-Item $key -Force | Out-Null }
Set-ItemProperty $key -Name AccessVBOM -Value 1 -Type DWord
if (Test-Path $WorkFolder) { Remove-Item $WorkFolder -Recurse -Force }
New-Item -ItemType Directory $WorkFolder | Out-Null
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Id

$job = Start-Job -ArgumentList $ModuleFolder, $Modules, $WorkFolder -ScriptBlock {
    param($modDir, $modules, $work)
    $out = @(); $xl = $null
    function Chk($name, $got, $expected) {
        $ok = (("$got").Trim() -eq "$expected")
        $script:out += ('{0,-4} {1}' -f $(if ($ok) { 'OK' } else { 'FOUT' }), $name) + $(if (-not $ok) { "  expected [$expected] got [$got]" } else { '' })
    }
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        foreach ($m in $modules) {
            $tmp = Join-Path $work ($m + '.bas')
            $text = [IO.File]::ReadAllText((Join-Path $modDir ($m + '.bas')), [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
            [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::GetEncoding(1252))
            $null = $wb.VBProject.VBComponents.Import($tmp)
        }
        $helper = $wb.VBProject.VBComponents.Add(1); $helper.Name = 'tstHulp'
        $helper.CodeModule.AddFromString(@"
Public Function TstExample() As String
    TstExample = CStr(SheetExists("Sheet1"))
End Function
"@)
        $run = { param($f) $xl.Run("'" + $wb.Name + "'!" + $f) }
        Chk 'SheetExists finds Sheet1' (& $run 'TstExample') 'True'
        $wb.Close($false)
    } catch { $out += 'ERROR: ' + $_.Exception.Message }
    finally { if ($xl) { try { $xl.Quit() } catch {} } }
    $out
}
if (Wait-Job $job -Timeout 120) { Receive-Job $job } else { 'TIMEOUT (a VBE or Excel dialog is probably open)'; Stop-Job $job }
Start-Sleep 2
foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) { if ($before -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } }
if ($null -eq $old) { Remove-ItemProperty $key -Name AccessVBOM -ErrorAction SilentlyContinue } else { Set-ItemProperty $key -Name AccessVBOM -Value $old -Type DWord }
