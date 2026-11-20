<#
.SYNOPSIS
  Uses the duct sizing template as an engineer would, in its own hidden Word (your open documents
  are left alone): a value typed, another way of sizing picked, a value that can't work. Logs ok or
  FAIL for each check to OutDir\steps.log.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 2)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$failures = 0
function Check([string]$what, [bool]$ok) { if (-not $ok) { $script:failures++ }; Step ("{0} {1}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $what) }
function Text { $script:doc.Content.Text }
function Recalculate { [void](Call $script:api 'Recalculate'); WaitForCalcpad }
function Pick([string]$variable, [string]$text) {
    $entries = (Tagged "calcpad-input:$variable").DropdownListEntries
    for ($k = 1; $k -le $entries.Count; $k++) { if ($entries.Item($k).Text -like "$text*") { $entries.Item($k).Select(); return } }
    throw "no '$text' in $variable"
}
# An equation's text in Word has paragraph marks in it (fractions): (?s) and a short reach after a label.
function Shows([string]$label, [string]$value) { (Text) -match ('(?s)' + [regex]::Escape($label) + '.{0,120}?=\s*' + [regex]::Escape($value) + '\s') }
try {
    if (-not (Open-Template $OutDir (Join-Path $PSScriptRoot "out\Dokimos Duct Sizing v$Version.dotx"))) { return }
    Check 'the template arrives calculated: 500 L/s at 1 Pa/m, 1190 m' ((Shows 'Diameter for this friction rate' '331.9') -and (Prop $api 'ErrorCount') -eq 0)
    Check 'automatic calculation is on' ((Call $api 'GetSettings') -notmatch 'auto=0')
    Check 'the calculation is locked, the inputs open' (CalculationLocked)

    # 1. Typing a value and leaving the field recalculates the sizes.
    $q = Tagged 'calcpad-input:q'
    $q.Range.Select()
    $word.Selection.TypeText('1000')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $doc.Paragraphs.Item(1).Range.Select()
    $ok = WaitUntil { (Shows 'Friction loss (Darcy' '1') -and -not (Shows 'Diameter for this friction rate' '331.9') }
    Step ("q = 1000 L/s, field left: {0} ({1:F1} s)" -f (Prop $api 'LastStatus'), $timer.Elapsed.TotalSeconds)
    Check 'leaving a changed field sizes the larger duct' ($ok -and (Prop $api 'ErrorCount') -eq 0)

    # 2. Sized by velocity instead.
    Pick 'P_s' 'Velocity'
    (Tagged 'calcpad-input:x_s').Range.Text = '5'
    Step "by velocity, 5 m/s: $(Recalculate)"
    Check 'a velocity gives the diameter' ((Shows 'Diameter for this velocity' '504.6') -and (Prop $api 'ErrorCount') -eq 0)

    # 3. A value that can't work says so, and is no calculation error.
    (Tagged 'calcpad-input:x_s').Range.Text = '0'
    Step "velocity 0: $(Recalculate)"
    Check 'a zero value is refused' ((Text) -match 'must be above 0' -and (Prop $api 'ErrorCount') -eq 0)

    [void](Call $doc 'Close' @(0))
    Step "done: $failures failure(s)"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
