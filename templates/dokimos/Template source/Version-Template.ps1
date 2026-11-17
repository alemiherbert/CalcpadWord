# Opens a document in a separate hidden Word, updates the contents and page fields, renders every
# page to <Name>-pageN.png, and saves it as a template. Writes pid.txt and steps.log to OutDir.
param([Parameter(Mandatory)][string]$OutDir, [string]$In = 'user\v3.docx', [string]$Out = 'user\Dokimos Design Calculation v3.dotx', [string]$Name = 'v3')
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
Add-Type -AssemblyName System.Drawing
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = Call $word.Documents 'Open' @([string](Join-Path $here $In), $false, $false, $false)
    Step "opened: $($doc.Sections.Count) sections, $($doc.ContentControls.Count) controls"
    $window = $doc.ActiveWindow; $window.View.Type = 3
    if ($doc.TablesOfContents.Count -gt 0) { [void](Call $doc.TablesOfContents.Item(1) 'Update'); Step "contents updated" }
    [void]$doc.ComputeStatistics(2)
    foreach ($sec in @($doc.Sections)) { foreach ($k in 1, 2, 3) { [void]$sec.Headers.Item($k).Range.Fields.Update(); [void]$sec.Footers.Item($k).Range.Fields.Update() } }
    for ($k = 1; $k -le $doc.Sections.Count; $k++) {
        $sec = $doc.Sections.Item($k)
        $footer = ($sec.Footers.Item(1).Range.Text + $sec.Footers.Item(2).Range.Text).Trim()
        Step ("section {0}: first-page header {1}, sheet header '{2}', footer '{3}'" -f $k, $sec.PageSetup.DifferentFirstPageHeaderFooter,
              (($sec.Headers.Item(1).Range.Text -replace '\s+', ' ').Trim() -replace '^(.{0,90}).*', '$1'), $footer)
    }
    $pages = $window.ActivePane.Pages
    Step "pages: $($pages.Count)"
    for ($n = 1; $n -le $pages.Count; $n++) {
        $emf = Join-Path $here ("user\{0}-page{1}.emf" -f $Name, $n); [IO.File]::WriteAllBytes($emf, [byte[]]$pages.Item($n).EnhMetaFileBits)
        $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1400, [int]($mf.Height * 1400.0 / $mf.Width))
        $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.SmoothingMode = 'HighQuality'; $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
        $bmp.Save([IO.Path]::ChangeExtension($emf, '.png'), [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
    }
    $target = [string](Join-Path $here $Out)
    [void](Call $doc 'SaveAs2' @($target, 14))
    Step "saved $target"
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
