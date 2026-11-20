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
    $template = [string](Join-Path $here 'user\Dokimos Design Calculation v4.dotx')
    $doc = Call $word.Documents 'Add' @($template)
    Step "new document from template: $($doc.Name), attached template: $($doc.AttachedTemplate.Name)"
    # Fill the project once, on the cover.
    $project = @(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Tag -eq 'dokimos:project' })[0]
    $project.Range.Text = 'Kampala Northern Bypass'
    $headerText = $doc.Sections.Item($doc.Sections.Count).Headers.Item(1).Range.Text
    Step ("header shows the project: " + ($headerText -match 'Kampala Northern Bypass'))
    # No guidance: the body holds only headings, each followed by an empty paragraph to type in.
    $body = @(); $p = $doc.Paragraphs.First
    while ($null -ne $p) { if ($body.Count -gt 0 -or ($p.OutlineLevel -eq 1 -and $p.Range.Text.Trim() -eq 'Scope')) { $body += $p }; $p = $p.Next(1) }
    $text = @($body | Where-Object { $_.OutlineLevel -gt 2 -and $_.Range.Text.Trim().Length -gt 0 })
    Step ("no guidance in the body ({0} headings): {1}" -f @($body | Where-Object { $_.OutlineLevel -le 2 }).Count, ($text.Count -eq 0))
    Step ("cover hints gone: " + ($doc.Content.Text -notmatch 'e\.g\.|Subject:'))
    # Type a Calcpad line under Geometry, and convert it.
    $geometry = @($body | Where-Object { $_.Range.Text.Trim() -eq 'Geometry' })[0]
    $p = $geometry.Next(1)
    $range = $p.Range; [void]$range.MoveEnd(1, -1); $range.Text = 'L = 6m'
    $p.Range.Select()
    Step ("Convert to Calcpad -> " + (Call $api 'Evaluate'))
    Step ("paragraph now: style=$($p.Style.NameLocal) maths=$($p.Range.OMaths.Count)")
    $window = $doc.ActiveWindow; $window.View.Type = 3
    $page = $window.ActivePane.Pages.Item($window.ActivePane.Pages.Count)
    $emf = Join-Path $here 'out\test-page3.emf'; [IO.File]::WriteAllBytes($emf, [byte[]]$page.EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1400, [int]($mf.Height * 1400.0 / $mf.Width))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save((Join-Path $here 'out\test-page3.png'), [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
