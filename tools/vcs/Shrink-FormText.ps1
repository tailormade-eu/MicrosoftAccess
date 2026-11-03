<#
.SYNOPSIS
  Shrinks the text of a form (or report) to the smallest set of lines for which Access still refuses to save it.

.DESCRIPTION
  Use it when Access imports an object but will not save it ("changes that are incompatible with the current database
  format", error 13122) and nothing in the text tells you why. It is delta debugging (ddmin): it removes chunks of lines, loads
  the rest with LoadFromText into a scratch .mdb of format 10 (Access 2002), opens the form in design view and closes it with
  save. A candidate counts as "still failing" only when the refusal is the 13122 message; text that no longer loads at all
  (a broken Begin/End pair) does not count. What is left is the smallest text that reproduces the refusal.

  Run it again on the result of your fix to find a second, independent cause: the first run finds only one. In the case
  that made us write it, the two causes were a layout-grid control (EmptyCell) and the layout properties (LayoutGroup,
  GroupTable, ColumnStart/End, RowEnd, paddings): the format-2002 mdb accepts neither, and removing only one of them was
  not enough.

  Needs a hidden Access (it holds Shift only to be safe; the scratch database has no startup form). About 450 tests in 40 to
  60 seconds for a form of 220 lines. Stop an Access that stays behind with Get-CimInstance (command line contains -Embedding).

.EXAMPLE
  .\Shrink-FormText.ps1 -FormFile D:epoorms\Login.form -Work $env:TEMP\shrink
#>
param(
    [Parameter(Mandatory)][string]$FormFile,
    [Parameter(Mandatory)][string]$Work,
    [int]$MaxTests = 450,
    [int]$MaxSeconds = 1200
)
# Delta debugging (ddmin): shrinks a form text file to the smallest set of lines for which Access still accepts the load
# but refuses to SAVE the form with "incompatible with the current database format" in a format-10 (Access 2002) .mdb.
New-Item -ItemType Directory $Work -Force | Out-Null
$db = Join-Path $Work 'ddmin.mdb'
if (Test-Path $db) { Remove-Item $db -Force }
$e = New-Object -ComObject DAO.DBEngine.120
$d = $e.CreateDatabase($db, ';LANGID=0x0409;CP=1252;COUNTRY=0', 64); $d.Close()

Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class KB { [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, int e); }'
$acc = New-Object -ComObject Access.Application; $acc.Visible = $false; $acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($db)
"file format: " + $acc.CurrentProject.FileFormat

$text = [IO.File]::ReadAllText($FormFile, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF) -replace "`r?`n", "`r`n"
$all = @($text -split "`r`n")
$file = Join-Path $Work 'probe.txt'
$cache = @{}
$tests = 0
$start = Get-Date

function Probe([string[]]$ls) {
    $key = ($ls -join "`n").GetHashCode().ToString() + ':' + $ls.Count
    if ($cache.ContainsKey($key)) { return $cache[$key] }
    $script:tests++
    [IO.File]::WriteAllText($file, ($ls -join "`r`n"), [Text.Encoding]::GetEncoding(1252))
    $r = 'PASS'
    try { $acc.DoCmd.DeleteObject(2, 'probe') } catch {}
    try {
        $acc.LoadFromText(2, 'probe', $file)
        try {
            $acc.DoCmd.OpenForm('probe', 1)
            $acc.DoCmd.Close(2, 'probe', 1)
        } catch {
            if ($_.Exception.Message -match 'incompatible with the current database format') { $r = 'FAIL' } else { $r = 'OTHER' }
            try { $acc.DoCmd.Close(2, 'probe', 2) } catch {}
        }
    } catch { $r = 'LOADFAIL' }
    $cache[$key] = $r
    return $r
}

$base = Probe $all
"full file: $base ($($all.Count) lines)"
if ($base -ne 'FAIL') { "the full file does not reproduce; stop"; $acc.CloseCurrentDatabase(); $acc.Quit(2); return }

$cur = $all
$n = 2
while ($cur.Count -ge 2 -and $tests -lt $MaxTests -and ((Get-Date) - $start).TotalSeconds -lt $MaxSeconds) {
    $size = [math]::Ceiling($cur.Count / $n)
    $reduced = $false
    for ($i = 0; $i -lt $n; $i++) {
        $from = $i * $size
        if ($from -ge $cur.Count) { break }
        $to = [math]::Min($cur.Count, $from + $size)
        $comp = @(); if ($from -gt 0) { $comp += $cur[0..($from - 1)] }; if ($to -lt $cur.Count) { $comp += $cur[$to..($cur.Count - 1)] }
        if ($comp.Count -eq 0) { continue }
        if ((Probe $comp) -eq 'FAIL') { $cur = $comp; $n = [math]::Max($n - 1, 2); $reduced = $true; "tests=$tests lines=$($cur.Count)"; break }
    }
    if (-not $reduced) {
        if ($n -ge $cur.Count) { break }
        $n = [math]::Min($cur.Count, $n * 2)
    }
}
"DONE after $tests tests, $([int]((Get-Date) - $start).TotalSeconds) s; minimal failing text has $($cur.Count) lines:"
"-----"
$cur
"-----"
$acc.CloseCurrentDatabase(); $acc.Quit(2)
