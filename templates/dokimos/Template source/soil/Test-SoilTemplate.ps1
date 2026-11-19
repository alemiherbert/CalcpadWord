<#
.SYNOPSIS
  Uses the soil stresses and settlement template as an engineer would, in its own hidden Word (your
  open documents are left alone). Logs ok or FAIL for each check to OutDir\steps.log.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 1)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$failures = 0
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Text { $script:doc.Content.Text }
function Recalculate { [void](Call $script:api 'Recalculate'); WaitForCalcpad }
# An equation's text in Word has paragraph marks in it (fractions): (?s) and a short reach after a label.
function Shows([string]$label, [string]$pattern) { (Text) -match ('(?s)' + [regex]::Escape($label) + '.{0,60}?=\s*' + $pattern + '\s*mm') }
try {
    if (-not (Open-Template $OutDir (Join-Path $PSScriptRoot "out\Dokimos Soil Stresses and Settlement v$Version.dotx"))) { return }
    Check 'the template arrives calculated: 10.23 mm at the centre' ((Shows 'At the centre' '10\.23') -and (Prop $api 'ErrorCount') -eq 0)
    Check 'automatic calculation is off' ((Call $api 'GetSettings') -match 'auto=0')
    $plots = 0; $shapes = $doc.InlineShapes
    for ($k = 1; $k -le $shapes.Count; $k++) { if ($shapes.Item($k).Title -eq 'Calcpad plot') { $plots++ } }
    Check "the plots and maps are there ($plots)" ($plots -eq 6)

    # 1. Typing a value and leaving the field is instant.
    $p2 = Tagged 'calcpad-input:p_2:kPa' # an input with a unit is tagged with it
    $p2.Range.Select()
    $word.Selection.TypeText('50')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $doc.Paragraphs.Item(1).Range.Select()
    for ($i = 0; $i -lt 100 -and (Prop $api 'LastStatus') -notmatch 'out of date'; $i++) { Start-Sleep -Milliseconds 50 }
    Check ("leaving a changed field marks the results out of date ({0:F2} s)" -f $timer.Elapsed.TotalSeconds) (((Prop $api 'LastStatus') -match 'out of date') -and $timer.Elapsed.TotalSeconds -lt 1)

    # 2. A uniform 50 kPa: the linear part of the old load added nothing at the centre, so it halves.
    $timer.Restart()
    Step ("p_2 = 50 kPa: {0} ({1:F1} s)" -f (Recalculate), $timer.Elapsed.TotalSeconds)
    Check 'a uniform 50 kPa halves the settlement at the centre (5.12 mm)' ((Shows 'At the centre' '5\.1[1-3]') -and (Prop $api 'ErrorCount') -eq 0)

    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
