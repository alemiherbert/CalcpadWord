# In a separate hidden Word: is the installed add-in connected, does it calculate, and where was
# .NET loaded from? Writes steps.log and pid.txt to OutDir. Never touches other Word processes.
param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$log = Join-Path $OutDir 'steps.log'; Remove-Item $log -ErrorAction SilentlyContinue
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
$wordPid = $new[0]
Set-Content (Join-Path $OutDir 'pid.txt') $wordPid
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $addIn = $word.COMAddIns.Item('CalcpadWord.Connect')
    Step "add-in connected: $($addIn.Connect)"
    $api = $addIn.Object
    $doc = $word.Documents.Add()
    $doc.Content.Text = 'a = 2 + 3'
    $doc.Paragraphs.Item(1).Range.Select()
    Step ("Evaluate -> " + (Call $api 'Evaluate'))
    for ($i = 0; $i -lt 60 -and (Prop $api 'IsWorking'); $i++) { Start-Sleep -Milliseconds 250 }
    Step ("status: " + (Prop $api 'LastStatus'))
    Step ("document: equations=$($doc.OMaths.Count), text='" + ($doc.Content.Text -replace '\s+', ' ').Trim() + "'")
    Step ("math fonts (+ installed): " + (Prop $api 'MathFontList'))
    Get-Process -Id $wordPid -Module | Where-Object { $_.ModuleName -match '^(hostfxr|hostpolicy|coreclr|clr|CalcpadWord.*)\.dll$' } |
        ForEach-Object { Step ("module " + $_.ModuleName + " <- " + $_.FileName) }
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally {
    try { [void](Call $word 'Quit' @(0)) } catch { }
    Start-Sleep -Seconds 3
    if (Get-Process -Id $wordPid -ErrorAction SilentlyContinue) { Stop-Process -Id $wordPid -Force; Step "stopped hidden Word $wordPid" }
}
