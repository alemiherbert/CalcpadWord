param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
function Docs() { (1..$word.Documents.Count | ForEach-Object { $word.Documents.Item($_).Name }) -join ',' }
function State([string]$when) {
    Step ("{0}: docs=[{1}] active={2} doc.OMathFontName={3} api.MathFont={4}" -f $when, (Docs), $word.ActiveDocument.Name, $doc.OMathFontName, (Prop $api 'MathFont'))
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $doc = $word.Documents.Add()
    $doc.Content.Text = "F = 120kN`rb = 0.3m`rh = 0.5m`r$([char]0x03C3) = F/(b*h)|MPa"
    $doc.Content.Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate'))
    State 'after Evaluate'
    $bits = $doc.Content.EnhMetaFileBits
    State 'after EnhMetaFileBits'
    Step ('SetMathFont -> ' + (Call $api 'SetMathFont' @('Latin Modern Math')))
    for ($i = 0; $i -lt 50 -and (Prop $api 'IsWorking'); $i++) { Start-Sleep -Milliseconds 100 }
    State 'after SetMathFont'
    $xml = $doc.Content.ContentControls.Item(1).Range.WordOpenXML
    $fonts = [regex]::Matches($xml, 'w:ascii="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    Step ("fonts in F = 120kN: " + ($fonts -join ', '))
    $m = [regex]::Match($xml, '<m:oMath>.{0,1500}')
    Step ("math xml: " + $m.Value)
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
