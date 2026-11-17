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
function Picture($range, [string]$name) {
    Add-Type -AssemblyName System.Drawing
    $emf = Join-Path $OutDir "$name.emf"
    [IO.File]::WriteAllBytes($emf, [byte[]]$range.EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf)
    $bmp = New-Object System.Drawing.Bitmap([int]($mf.Width * 1.5), [int]($mf.Height * 1.5))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save((Join-Path $OutDir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $doc = $word.Documents.Add()
    $doc.Content.Text = "F = 120kN`rb = 0.3m`rh = 0.5m`r$([char]0x03C3) = F/(b*h)|MPa"
    $doc.Content.Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate'))
    State 'after Evaluate'
    Step ('fonts: ' + (Prop $api 'MathFontList'))
    State 'after MathFontList'
    try { [void](Call $api 'SetMathFont' @('Minion Math')) } catch { Step ('Minion: ' + $_.Exception.InnerException.Message) }
    State 'after Minion'
    foreach ($family in 'Euler Math', 'STIX Two Math', 'Cambria Math') {
        Step ("SetMathFont $family -> " + (Call $api 'SetMathFont' @($family)))
        for ($i = 0; $i -lt 300 -and (Prop $api 'IsWorking'); $i++) { if ($i % 5 -eq 0) { State "  waiting $i" }; Start-Sleep -Milliseconds 100 }
        State "after $family"
        Picture $doc.Content ('font-' + ($family -replace ' ', '-'))
        State "after picture"
    }
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
