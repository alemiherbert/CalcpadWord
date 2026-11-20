<#
.SYNOPSIS
  Uses the beam template as a designer would, in its own hidden Word (your open documents are
  left alone): inputs typed, picked and mistyped, results checked. Logs ok or FAIL for each
  check to OutDir\steps.log, and draws the results page to out\test-results.png.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 3)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$failures = 0
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Text { $script:doc.Content.Text }
function Recalculate { [void](Call $script:api 'Recalculate'); WaitForCalcpad }
try {
    if (-not (Open-Template $OutDir (Join-Path $PSScriptRoot "out\Dokimos RC Beam Design (EC2) v$Version.dotx"))) { return }
    Check 'the template arrives calculated, passing' ((Text) -match 'The beam passes every check')

    # 1. Calculation happens on Recalculate only. Leaving a changed field just marks the
    #    results out of date, at once; Recalculate updates them.
    Check 'automatic calculation is off' ((Call $api 'GetSettings') -match 'auto=0')
    (Tagged 'calcpad-input:h').Range.Select()
    $word.Selection.TypeText('350')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $doc.Paragraphs.Item(1).Range.Select() # leave the field
    for ($i = 0; $i -lt 100 -and (Prop $api 'LastStatus') -notmatch 'out of date'; $i++) { Start-Sleep -Milliseconds 50 }
    Step ("after typing h = 350 and leaving: {0} ({1:F2} s)" -f (Prop $api 'LastStatus'), $timer.Elapsed.TotalSeconds)
    Check 'leaving a changed field marks the results out of date, within a second' (((Prop $api 'LastStatus') -match 'out of date') -and $timer.Elapsed.TotalSeconds -lt 1)
    Check 'and calculates nothing' ((Text) -notmatch '350\s*[−-]\s*30\s*[−-]\s*8')
    [void](Recalculate)
    Check 'Recalculate updates the results' ((Text) -match '350\s*[−-]\s*30\s*[−-]\s*8')

    # 2. A shallow, long beam fails, and the summary says so.
    (Tagged 'calcpad-input:L').Range.Text = '8'
    (Tagged 'calcpad-input:h').Range.Text = '300'
    Step "h = 300, L = 8: $(Recalculate)"
    Check 'the summary reports a failure' ((Text) -match 'The beam fails at least one check')
    Check 'a failing check says FAIL' ((Text) -match 'FAIL')

    # 3. Pick a concrete class from the drop-down: its value is f_ck.
    $entries = (Tagged 'calcpad-input:f_ck').DropdownListEntries
    for ($k = 1; $k -le $entries.Count; $k++) { if ($entries.Item($k).Text -eq 'C40/50') { $entries.Item($k).Select() } }
    Step "C40/50 picked: $(Recalculate)"
    $next = (Tagged 'calcpad-input:f_ck').Range.Paragraphs.Item(1).Next(1).Range.Text
    Check 'the class sets f_ck = 40' ($next -match '40')

    # 4. A value that isn't a number is flagged, and the rest still calculates.
    (Tagged 'calcpad-input:b').Range.Text = '30O'
    $status = Recalculate
    Step "b = 30O: $status"
    Check 'a bad value is reported' ($status -match 'b: enter a number')
    Check 'it counts as an error' ((Prop $api 'ErrorCount') -ge 1)
    (Tagged 'calcpad-input:b').Range.Text = '300'
    [void](Recalculate)
    Check 'fixing it clears the error' ((Prop $api 'ErrorCount') -eq 0)

    # Back to a passing beam, and a picture of the results.
    (Tagged 'calcpad-input:L').Range.Text = '6'
    (Tagged 'calcpad-input:h').Range.Text = '500'
    [void](Recalculate)
    Check 'back to passing' ((Text) -match 'The beam passes every check')
    Save-PagePicture $doc.ActiveWindow.ActivePane.Pages.Count (Join-Path $PSScriptRoot 'out\test-results.png')
    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
