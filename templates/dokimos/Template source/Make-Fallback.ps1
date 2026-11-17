# Has Word make the PNG fallback for the corrected pattern SVG: inserts it into a new document
# at the user's size and saves it; the PNG is then taken from that package.
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
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $svg = [string](Join-Path $here 'out\cover-pattern.svg')
    $shape = $doc.Content.InlineShapes.AddPicture($svg, $false, $true)
    Step ("inserted: {0} x {1} pt" -f $shape.Width, $shape.Height)
    $target = [string](Join-Path $here 'user\fallback.docx')
    [void](Call $doc 'SaveAs2' @($target, 16))
    Step "saved $target"
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
