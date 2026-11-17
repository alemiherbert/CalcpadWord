param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Picture($range, [string]$name) {
    Add-Type -AssemblyName System.Drawing
    $emf = Join-Path $OutDir "$name.emf"
    [IO.File]::WriteAllBytes($emf, [byte[]]$range.EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf)
    $bmp = New-Object System.Drawing.Bitmap([int]($mf.Width * 1.0), [int]($mf.Height * 1.0))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White)
    $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save((Join-Path $OutDir "$name.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
}
function Body($doc) { [regex]::Match($doc.Content.WordOpenXML, '<w:body>.*</w:body>', 'Singleline').Value -replace ' w:rsid\w*="[^"]*"', '' }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    foreach ($method in 'shading') {
        $doc = $word.Documents.Add()
        $doc.Content.Text = "F = 120kN`rThe stress is {= s = F/(0.3m*0.5m)|MPa} here.`r{=`r'Plot`r`$Plot{sin(x) @ x = 0 : 6}`r}`rSmall: {= x = 0.0000000057m}"
        $doc.Paragraphs.Item(1).Range.Select()
        Step ("$method Evaluate -> " + (Call $api 'Evaluate'))
        for ($w = 0; $w -lt 100 -and [System.__ComObject].InvokeMember('IsWorking', [Reflection.BindingFlags]::GetProperty, $null, $api, $null); $w++) { Start-Sleep -Milliseconds 200 }
        Step ("  status: " + [System.__ComObject].InvokeMember('LastStatus', [Reflection.BindingFlags]::GetProperty, $null, $api, $null))
        $original = Body $doc
        Picture $doc.Content "$method-0-before"
        for ($i = 1; $i -le $doc.ContentControls.Count; $i++) {
            $cc = $doc.ContentControls.Item($i)
            $cc.LockContents = $false
            if ($method -eq 'shading') { $cc.Range.Font.Shading.BackgroundPatternColor = 15132390 }   # wdColorGray10
            else { $cc.Range.Font.Fill.Transparency = 0.6 }
            for ($j = 1; $j -le $cc.Range.InlineShapes.Count; $j++) { $cc.Range.InlineShapes.Item($j).PictureFormat.ColorType = 4 }  # msoPictureWatermark
            $cc.LockContents = $true
        }
        Picture $doc.Content "$method-1-stale"
        for ($i = 1; $i -le $doc.ContentControls.Count; $i++) {
            $cc = $doc.ContentControls.Item($i)
            $cc.LockContents = $false
            if ($method -eq 'shading') { $cc.Range.Font.Shading.BackgroundPatternColor = -16777216 }  # wdColorAutomatic
            else { $cc.Range.Font.Fill.Transparency = 0 }
            for ($j = 1; $j -le $cc.Range.InlineShapes.Count; $j++) { $cc.Range.InlineShapes.Item($j).PictureFormat.ColorType = 1 }  # automatic
            $cc.LockContents = $true
        }
        Picture $doc.Content "$method-2-restored"
        $restored = Body $doc
        Step ("$method restored identical XML: {0} (lengths {1} / {2})" -f ($original -eq $restored), $original.Length, $restored.Length)
        if ($original -ne $restored) {
            $k = 0; while ($k -lt [Math]::Min($original.Length, $restored.Length) -and $original[$k] -eq $restored[$k]) { $k++ }
            Step ("  first difference at $k : was [" + $original.Substring([Math]::Max(0, $k - 80), [Math]::Min(300, $original.Length - [Math]::Max(0, $k - 80))) + "]")
            Step ("  now [" + $restored.Substring([Math]::Max(0, $k - 80), [Math]::Min(300, $restored.Length - [Math]::Max(0, $k - 80))) + "]")
        }
        [void](Call $doc 'Close' @(0))
    }
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
