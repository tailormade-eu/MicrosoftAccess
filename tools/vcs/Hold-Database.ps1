<#
.SYNOPSIS
  Keeps a database open in a hidden Access, so the msaccess-vcs-mcp tools (for example vcs_compile_vba) can work on it.

.DESCRIPTION
  vcs_compile_vba needs a database that is already open in an Access that the add-in can reach. Start this script in
  its own window (Start-Process pwsh -WindowStyle Hidden ...), wait about 30 s, call the MCP tool, then create the file
  stop.flag next to the database: the script closes the database, quits Access and removes the flag.

  A compile check that cannot fail proves nothing: always run it once on a module you broke on purpose. The command
  RunCommand(126) (compile and save all modules) does NOT report a compile error in a hidden Access.

.EXAMPLE
  Start-Process pwsh -ArgumentList '-NoProfile','-File','Hold-Database.ps1','C:\temp\MyApp.accdb' -WindowStyle Hidden
  # ... vcs_compile_vba("C:\temp\MyApp.accdb") ...
  New-Item C:\temp\stop.flag -Force | Out-Null
#>
param([Parameter(Mandatory)] [string] $Db, [int] $MaxMinutes = 10)
$stop = Join-Path (Split-Path $Db) 'stop.flag'
if (Test-Path $stop) { Remove-Item $stop -Force }
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($Db)
$end = (Get-Date).AddMinutes($MaxMinutes)
while (-not (Test-Path $stop) -and (Get-Date) -lt $end) { Start-Sleep -Seconds 2 }
try { $acc.CloseCurrentDatabase() } catch {}
try { $acc.Quit(2) } catch {}
if (Test-Path $stop) { Remove-Item $stop -Force }
