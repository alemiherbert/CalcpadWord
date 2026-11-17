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
    $doc = Call $word.Documents 'Open' @([string](Join-Path $here 'out\dokimos-template.docx'), $false, $false, $false)
    Step "opened; sections $($doc.Sections.Count)"
    for ($k = 1; $k -le $doc.Sections.Count; $k++) {
        $sec = $doc.Sections.Item($k)
        $h = $sec.Headers.Item(1)
        Step "section $k header: exists=$($h.Exists) linked=$($h.LinkToPrevious) chars=$($h.Range.Characters.Count)"
    }
    Step "pages: $($doc.ComputeStatistics(2))"
    $h = $doc.Sections.Item(2).Headers.Item(1)
    Step "section 2 header text: $(($h.Range.Text -replace '\s+', ' ').Substring(0, 60))"
    $paras = $h.Range.Paragraphs
    Step "section 2 header paragraphs: $($paras.Count)"
    $first = $paras.Item(1).Range
    Step "first paragraph text: [$($first.Text)] in table: $($first.Information(12))"
    $logo = [string](Resolve-Path (Join-Path $here '..\logo-download\logo.svg')).Path
    $r = $first.Duplicate
    $i = $first.Text.IndexOf('{{LOGO_SHEET}}')
    Step "marker at $i"
    $r.SetRange($first.Start + $i, $first.Start + $i + 14)
    $r.Text = ''
    Step "marker removed"
    $shape = $h.Range.InlineShapes.AddPicture($logo, $false, $true, $r)
    Step "picture added: $($shape.Width) x $($shape.Height)"
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
