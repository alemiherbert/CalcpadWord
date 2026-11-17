param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    for ($w = 0; $w -lt 100 -and (Prop $api 'IsWorking'); $w++) { Start-Sleep -Milliseconds 200 }
    Step 'warmed up'
    foreach ($round in 1, 2, 3) {
        $doc = $word.Documents.Add()
        $doc.Content.Text = (@('L = 6m', '{=', '"Beam check', "'The span is 'L", 'M = 3*L', '#if M > 10m', "'Long beam", '#else', "'Short beam", '#end if', '#for i = 1 : 3', "'Step 'i", '#loop', '$Plot{x^2 @ x = 0 : 2}', '}', 'Moment from the block: {=#val M}') -join "`r")
        $doc.Paragraphs.Item(1).Range.Select()
        Step "round $round Evaluate -> $(Call $api 'Evaluate')"
        for ($w = 0; $w -lt 100 -and (Prop $api 'IsWorking'); $w++) { Start-Sleep -Milliseconds 100 }
        Step "  done: $(Prop $api 'LastStatus')"
        Step "  controls: $($doc.ContentControls.Count)"
        for ($i = 1; $i -le $doc.ContentControls.Count; $i++) { $cc = $doc.ContentControls.Item($i); Step "  [$i] $($cc.Tag.Substring(0, [Math]::Min(40, $cc.Tag.Length))) | $($cc.Title)" }
        [void](Call $doc 'Close' @(0))
        Step '  closed'
    }
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
