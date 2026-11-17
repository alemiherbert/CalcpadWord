param([string]$Template, [string]$Png)
function Call($obj, [string]$m, [object[]]$a = @()) { [System.__ComObject].InvokeMember($m, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $a) }
Add-Type -AssemblyName System.Drawing
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { 'not separate'; return }
try {
  $word.Visible = $false; $word.DisplayAlerts = 0
  $doc = Call $word.Documents 'Add' @($Template)
  $w = $doc.ActiveWindow; $w.View.Type = 3
  $page = $w.ActivePane.Pages.Item(4)
  $emf = [IO.Path]::ChangeExtension($Png, '.emf'); [IO.File]::WriteAllBytes($emf, [byte[]]$page.EnhMetaFileBits)
  $mf = New-Object System.Drawing.Imaging.Metafile($emf); $W = 3600; $H = [int]($mf.Height * $W / $mf.Width)
  $bmp = New-Object System.Drawing.Bitmap($W, $H); $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.SmoothingMode = 'HighQuality'; $g.TextRenderingHint = 'AntiAliasGridFit'
  $g.DrawImage($mf, 0, 0, $W, $H)
  $crop = $bmp.Clone((New-Object System.Drawing.Rectangle(0, [int]($H*0.47), [int]($W*0.7), [int]($H*0.13))), $bmp.PixelFormat); $crop.Save($Png)
  $g.Dispose(); $bmp.Dispose(); $crop.Dispose(); $mf.Dispose(); Remove-Item $emf
  [void](Call $doc 'Close' @(0))
} finally { try { [void](Call $word 'Quit' @(0)) } catch {}; Start-Sleep 2; if (Get-Process -Id $new[0] -ErrorAction SilentlyContinue) { Stop-Process -Id $new[0] -Force } }
