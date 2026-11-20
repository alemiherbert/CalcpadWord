<#
.SYNOPSIS
  Uses the psychrometric chart template as an engineer would, in its own hidden Word (your open
  documents are left alone): values typed, properties picked, a pair that can't work. Logs ok or
  FAIL for each check to OutDir\steps.log, and draws the chart page to out\test-chart.png.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 3)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$failures = 0
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Text { $script:doc.Content.Text }
function Recalculate { [void](Call $script:api 'Recalculate'); WaitForCalcpad }
function Pick([string]$variable, [string]$text) {
    $entries = (Tagged "calcpad-input:$variable").DropdownListEntries
    for ($k = 1; $k -le $entries.Count; $k++) { if ($entries.Item($k).Text -eq $text) { $entries.Item($k).Select(); return } }
    throw "no '$text' in $variable"
}
try {
    if (-not (Open-Template $OutDir (Join-Path $PSScriptRoot "out\Dokimos Psychrometric Chart v$Version.dotx"))) { return }
    Check 'the template arrives calculated: 25 °C and 60 % at 1190 m' ((Text) -match 'Relative humidity' -and (Prop $api 'ErrorCount') -eq 0)
    Check 'automatic calculation is off' ((Call $api 'GetSettings') -match 'auto=0')

    # 1. Typing a value and leaving the field is instant; Recalculate updates the results.
    $x2 = Tagged 'calcpad-input:x_2'
    $x2.Range.Select()
    $word.Selection.TypeText('40')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $doc.Paragraphs.Item(1).Range.Select()
    for ($i = 0; $i -lt 100 -and (Prop $api 'LastStatus') -notmatch 'out of date'; $i++) { Start-Sleep -Milliseconds 50 }
    Check ("leaving a changed field marks the results out of date ({0:F2} s)" -f $timer.Elapsed.TotalSeconds) (((Prop $api 'LastStatus') -match 'out of date') -and $timer.Elapsed.TotalSeconds -lt 1)
    $timer.Restart()
    Step ("φ = 40 %: {0} ({1:F1} s)" -f (Recalculate), $timer.Elapsed.TotalSeconds)
    # An equation's text in Word has paragraph marks in it (fractions): (?s) and a short reach.
    Check 'Recalculate finds the air' ((Text) -match '(?s)Relative humidity:.{0,80}?=\s*40\s*%' -and (Prop $api 'ErrorCount') -eq 0)

    # 2. Another pair: dry bulb and wet bulb.
    Pick 'P_2' 'Wet-bulb temperature, °C'
    $x2.Range.Text = '18'
    Step "t = 25 °C, t_wb = 18 °C: $(Recalculate)"
    Check 'a dry and a wet bulb give the air' ((Text) -match '(?s)psychrometric equation\):.{0,40}?=\s*18\s*°C' -and (Prop $api 'ErrorCount') -eq 0)

    # 3. A pair that can't fix the air says so, and calculates no chart.
    Pick 'P_2' 'Dry-bulb temperature, °C'
    Step "the same property twice: $(Recalculate)"
    Check 'the same property twice is refused' ((Text) -match 'Choose two different known properties' -and (Text) -match 'No chart')
    Check 'which is no calculation error' ((Prop $api 'ErrorCount') -eq 0)

    # Back to the defaults, and a picture of the chart.
    Pick 'P_2' 'Relative humidity, %'
    $x2.Range.Text = '60'
    [void](Recalculate)
    Check 'back to the default air' ((Text) -match '(?s)Relative humidity:.{0,80}?=\s*60\s*%')
    $chart = $doc.InlineShapes
    for ($k = 1; $k -le $chart.Count; $k++) {
        if ($chart.Item($k).Title -eq 'Calcpad plot') { Save-PagePicture $chart.Item($k).Range.Information(3) (Join-Path $PSScriptRoot 'out\test-chart.png'); break }
    }
    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
