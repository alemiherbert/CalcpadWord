<#
.SYNOPSIS
  Uses the seepage template as an engineer would, in its own hidden Word (your open documents are
  left alone): a deeper wall, a shallow one that fails, a value typed wrongly. Logs ok or FAIL for
  each check to OutDir\steps.log, and draws the flow net page to out\test-flownet.png.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 1)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$failures = 0
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Text { $script:doc.Content.Text }
function Recalculate { [void](Call $script:api 'Recalculate'); WaitForCalcpad }
# N_f is the last value on the line "Flow channels (…): N_f = …"; an equation's text in Word has paragraph marks in it.
function Channels([string]$pattern) { (Text) -match ('(?s)Flow channels \(.{0,200}?=\s*' + $pattern + '\s') }
try {
    if (-not (Open-Template $OutDir (Join-Path $PSScriptRoot "out\Dokimos Seepage by Finite Differences v$Version.dotx"))) { return }
    Check 'the template arrives calculated: 4.31 flow channels for 12 drops (Craig: 4.3)' ((Channels '4\.31') -and (Prop $api 'ErrorCount') -eq 0)
    Check 'automatic calculation is on' ((Call $api 'GetSettings') -notmatch 'auto=0')
    Check 'the calculation is locked, the inputs open' (CalculationLocked)
    Check 'both checks pass' (([regex]::Matches((Text), 'PASS -')).Count -eq 2)
    # The drawings, in order: the section, the flow net, the water pressure on the wall.
    Check "the three drawings are there ($($doc.InlineShapes.Count))" ($doc.InlineShapes.Count -eq 3)

    # 1. A deeper wall, typed and left: the flow net recalculates, with fewer flow channels.
    $d = Tagged 'calcpad-input:d:m'
    $d.Range.Select()
    $word.Selection.TypeText('7')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $doc.Paragraphs.Item(1).Range.Select()
    $ok = WaitUntil { -not (Channels '4\.31') -and (Channels '\d\.\d+') } 60
    Step ("d = 7 m, field left: {0} ({1:F1} s)" -f (Prop $api 'LastStatus'), $timer.Elapsed.TotalSeconds)
    Check 'leaving a changed field recalculates the flow net' ($ok -and (Prop $api 'ErrorCount') -eq 0)

    # 2. A wall 1 m deep: the exit gradient is too high, and the template says so (no calculation error).
    (Tagged 'calcpad-input:d:m').Range.Text = '1'
    Step "d = 1 m: $(Recalculate)"
    Check 'a shallow wall fails at the exit' ((Text) -match 'FAIL - Exit gradient too high' -and (Prop $api 'ErrorCount') -eq 0)

    # 3. A permeability typed wrongly is reported; E notation is accepted.
    (Tagged 'calcpad-input:k_1:m/s').Range.Text = 'le-5' # a letter l for the 1
    $status = Recalculate
    Check "a bad value is reported ($status)" ($status -match 'k_1: enter a number')
    (Tagged 'calcpad-input:k_1:m/s').Range.Text = '2.5E-5'
    (Tagged 'calcpad-input:d:m').Range.Text = '6'
    [void](Recalculate)
    Check 'fixing it clears the error; 2.5E-5 m/s is read' ((Prop $api 'ErrorCount') -eq 0 -and (Channels '4\.31'))

    Save-PagePicture $doc.InlineShapes.Item(2).Range.Information(3) (Join-Path $PSScriptRoot 'out\test-flownet.png')
    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
