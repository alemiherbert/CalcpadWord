param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $doc = Call $word.Documents 'Add' @([string](Join-Path $here 'out\Dokimos Design Calculation.dotx'))
    Step "pages: $($doc.ComputeStatistics(2))"
    $first = $doc.Sections.Item(1).Footers.Item(2).Range
    [void]$first.Fields.Update()
    Step ("cover footer: " + ($first.Text -replace '\s+', ' '))
    Step ("sheet footer: " + ($doc.Sections.Item(1).Footers.Item(1).Range.Text -replace '\s+', ' '))
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
