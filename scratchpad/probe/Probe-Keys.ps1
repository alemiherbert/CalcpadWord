param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()   # key bindings are resolved in a document context
    $ctrl = 512; $shift = 256; $alt = 1024
    $keys = [ordered]@{
        'F5' = 116; 'Enter' = 13; 'E' = 69; 'R' = 82; 'D' = 68; 'Q' = 81; 'J' = 74; 'Y' = 89; 'K' = 75; 'B' = 66
        'F9' = 120; 'F8' = 119; 'F12' = 123; '=' = 187
    }
    $mods = [ordered]@{
        'Ctrl+' = $ctrl; 'Alt+' = $alt; 'Shift+' = $shift; 'Ctrl+Shift+' = $ctrl + $shift; 'Ctrl+Alt+' = $ctrl + $alt
        'Alt+Shift+' = $alt + $shift; 'Ctrl+Alt+Shift+' = $ctrl + $alt + $shift
    }
    foreach ($m in $mods.Keys) {
        $line = foreach ($k in $keys.Keys) {
            $cmd = $word.FindKey($mods[$m] + $keys[$k]).Command
            "{0}{1}={2}" -f $m, $k, $(if ($cmd) { $cmd } else { '(free)' })
        }
        Step ($line -join ' | ')
    }
    foreach ($k in 'F5', 'F9') { Step ("{0}={1}" -f $k, $word.FindKey($keys[$k]).Command) }
    $doc.Close([ref]0)
}
catch { Step ("FAILED: " + $_.Exception.Message) }
finally { try { $word.Quit([ref]0) } catch { Step "quit failed: $($_.Exception.Message)" } }
