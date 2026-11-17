<#
.SYNOPSIS
  Uses the beam template as a designer would, in its own hidden Word (your open documents are
  left alone): a new document from the template, inputs typed, picked and mistyped, results checked.
  Writes steps.log and pid.txt to OutDir, and a picture of the results page to out\test-results.png.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 1)
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$template = [string](Join-Path $here "out\Dokimos RC Beam Design (EC2) v$Version.dotx")
New-Item -ItemType Directory -Force $OutDir | Out-Null
$log = Join-Path $OutDir 'steps.log'; Remove-Item $log -ErrorAction SilentlyContinue
$failures = 0
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Call($obj, [string]$method, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$property) { [System.__ComObject].InvokeMember($property, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
function Tagged([string]$tag) {
    $all = $doc.ContentControls
    for ($i = 1; $i -le $all.Count; $i++) { $c = $all.Item($i); if ($c.Tag -eq $tag) { return $c } }
    throw "No content control tagged $tag"
}
function WaitForCalcpad { for ($i = 0; $i -lt 480 -and (Prop $api 'IsWorking'); $i++) { Start-Sleep -Milliseconds 250 }; Prop $api 'LastStatus' }
function Text { $doc.Content.Text }
Add-Type -AssemblyName System.Drawing

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
$wordPid = $new[0]
Set-Content (Join-Path $OutDir 'pid.txt') $wordPid
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $doc = Call $word.Documents 'Add' @($template)
    Step "new document from $(Split-Path $template -Leaf)"
    Check 'the template arrives calculated, passing' ((Text) -match 'The beam passes every check')

    # 1. Type a new depth and leave the field: that alone recalculates (ContentControlOnExit).
    $h = Tagged 'calcpad-input:h'
    $h.Range.Select()
    $word.Selection.TypeText('350')
    $doc.Paragraphs.Item(1).Range.Select() # leave the field
    Start-Sleep -Milliseconds 800
    $status = WaitForCalcpad
    Step "after typing h = 350 and leaving: $status"
    Check 'leaving a changed field recalculates' ((Text) -match '350\s*[−-]\s*30\s*[−-]\s*8')

    # 2. A shallow, long beam fails, and the summary says so.
    (Tagged 'calcpad-input:L').Range.Text = '8'
    (Tagged 'calcpad-input:h').Range.Text = '300'
    $status = Call $api 'Recalculate'; $status = WaitForCalcpad
    Step "h = 300, L = 8: $status"
    Check 'the summary reports a failure' ((Text) -match 'The beam fails at least one check')
    Check 'a failing check says FAIL' ((Text) -match 'FAIL')

    # 3. Pick a concrete class from the drop-down: its value is f_ck.
    $fck = Tagged 'calcpad-input:f_ck'
    $entries = $fck.DropdownListEntries
    for ($k = 1; $k -le $entries.Count; $k++) { if ($entries.Item($k).Text -eq 'C40/50') { $entries.Item($k).Select() } }
    $status = Call $api 'Recalculate'; $status = WaitForCalcpad
    Step "C40/50 picked: $status"
    $paragraph = (Tagged 'calcpad-input:f_ck').Range.Paragraphs.Item(1).Next(1).Range.Text
    Step "next line: $($paragraph.Trim())"
    Check 'the class sets f_ck = 40' ($paragraph -match '40')

    # 4. A value that isn't a number is flagged, and the rest still calculates.
    (Tagged 'calcpad-input:b').Range.Text = '30O'
    $status = Call $api 'Recalculate'; $status = WaitForCalcpad
    Step "b = 30O: $status"
    Check 'a bad value is reported' ($status -match 'b: enter a number')
    Check 'it counts as an error' ((Prop $api 'ErrorCount') -ge 1)
    (Tagged 'calcpad-input:b').Range.Text = '300'
    $status = Call $api 'Recalculate'; $status = WaitForCalcpad
    Check 'fixing it clears the error' ((Prop $api 'ErrorCount') -eq 0)

    # Back to a passing beam, and a picture of the results.
    (Tagged 'calcpad-input:L').Range.Text = '6'
    (Tagged 'calcpad-input:h').Range.Text = '500'
    $status = Call $api 'Recalculate'; $status = WaitForCalcpad
    Check 'back to passing' ((Text) -match 'The beam passes every check')
    $window = $doc.ActiveWindow; $window.View.Type = 3
    $pages = $window.ActivePane.Pages
    $emf = Join-Path $here 'out\test-results.emf'; [IO.File]::WriteAllBytes($emf, [byte[]]$pages.Item($pages.Count).EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1200, [int]($mf.Height * 1200.0 / $mf.Width))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save([IO.Path]::ChangeExtension($emf, '.png'), [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally {
    try { [void](Call $word 'Quit' @(0)) } catch { }
    Start-Sleep -Seconds 3
    if (Get-Process -Id $wordPid -ErrorAction SilentlyContinue) { Stop-Process -Id $wordPid -Force; Step "stopped hidden Word $wordPid" }
}
