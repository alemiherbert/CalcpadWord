param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
Add-Type -AssemblyName System.Drawing
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $template = [string](Join-Path $here 'user\Dokimos Design Calculation v3.dotx')
    $doc = Call $word.Documents 'Add' @($template)
    Step "new document from template: $($doc.Name), attached template: $($doc.AttachedTemplate.Name)"
    # Fill the project once, on the cover.
    $project = @(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Tag -eq 'dokimos:project' })[0]
    $project.Range.Text = 'Kampala Northern Bypass'
    $headerText = $doc.Sections.Item($doc.Sections.Count).Headers.Item(1).Range.Text
    Step ("header shows the project: " + ($headerText -match 'Kampala Northern Bypass'))
    # Type a Calcpad line where the Geometry guidance is, and evaluate it.
    $guide = $null
    for ($i = 1; $i -le $doc.ContentControls.Count; $i++) { $c = $doc.ContentControls.Item($i); if ($c.Range.Text -like 'Dimensions and section properties*') { $guide = $c } }
    $range = $guide.Range
    $range.Text = 'L = 6m'
    Step ("guidance removed on typing: " + (@(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Range.Text -like 'Dimensions and section*' }).Count -eq 0))
    $p = @(1..$doc.Paragraphs.Count | ForEach-Object { $doc.Paragraphs.Item($_) } | Where-Object { $_.Range.Text -match '^L = 6m' })[0]
    $p.Range.Select()
    Step ("Evaluate -> " + (Call $api 'Evaluate'))
    Step ("paragraph now: style=$($p.Style.NameLocal) maths=$($p.Range.OMaths.Count)")
    $window = $doc.ActiveWindow; $window.View.Type = 3
    $page = $window.ActivePane.Pages.Item(3)
    $emf = Join-Path $here 'out\test-page3.emf'; [IO.File]::WriteAllBytes($emf, [byte[]]$page.EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1400, [int]($mf.Height * 1400.0 / $mf.Width))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save((Join-Path $here 'out\test-page3.png'), [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
