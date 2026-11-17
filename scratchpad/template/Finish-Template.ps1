# Finishes the generated packages in a separate hidden Word: logo into the headers, contents
# updated, page pictures rendered, saved as .dotx / .docx. Writes pid.txt and steps.log to OutDir.
param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$logo = (Resolve-Path (Join-Path $here '..\logo-download\logo.svg')).Path
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
Add-Type -AssemblyName System.Drawing
function PagePicture($page, [string]$png) {
    $emf = [IO.Path]::ChangeExtension($png, '.emf')
    [IO.File]::WriteAllBytes($emf, [byte[]]$page.EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf)
    $scale = 1400.0 / $mf.Width
    $bmp = New-Object System.Drawing.Bitmap(1400, [int]($mf.Height * $scale))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White)
    $g.SmoothingMode = 'HighQuality'; $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
}
function PlaceLogo($range, [string]$marker, [double]$points) {
    $paragraphs = $range.Paragraphs
    for ($n = 1; $n -le $paragraphs.Count; $n++) {
        $para = $paragraphs.Item($n)   # by index: enumerating a header's paragraphs crashed Word
        $text = $para.Range.Text
        $i = $text.IndexOf($marker)
        if ($i -lt 0) { continue }
        $r = $para.Range.Duplicate
        $r.SetRange($para.Range.Start + $i, $para.Range.Start + $i + $marker.Length)
        $r.Text = ''
        # The size comes from the SVG itself: resizing the picture afterwards made Word crash in a
        # multi-section document.
        $sized = Join-Path $here ("out\logo-{0}pt.svg" -f $points)
        $svg = [IO.File]::ReadAllText($logo) -replace 'width="100%" height="100%"', ('width="{0}pt" height="{0}pt"' -f $points)
        [IO.File]::WriteAllText($sized, $svg, [Text.UTF8Encoding]::new($false))
        $shape = $range.InlineShapes.AddPicture([string]$sized, $false, $true, $r)
        return $true
    }
    return $false
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($job in @(@{ In = 'dokimos-template.docx'; Out = 'Dokimos Design Calculation.dotx'; Format = 14; Pics = 'template' },
                       @{ In = 'dokimos-example.docx'; Out = 'Soil Stress Deformation (Dokimos).docx'; Format = 16; Pics = 'example' })) {
        $path = Join-Path $here ('out\' + $job.In)
        $doc = Call $word.Documents 'Open' @([string]$path, $false, $false, $false)
        Step "opened $($job.In): $($doc.Paragraphs.Count) paragraphs, $($doc.ContentControls.Count) controls"
        $section = $doc.Sections.Item(1)
        Step "  sections: $($doc.Sections.Count)"
        # Sections: 1 the cover (no header or footer), 2 the contents, 3 the body; 2 and 3 have the sheet header.
        $sheets = @()
        for ($k = 2; $k -le $doc.Sections.Count; $k++) {
            Step "  placing the sheet logo in section $k"
            $sheets += PlaceLogo $doc.Sections.Item($k).Headers.Item(1).Range '{{LOGO_SHEET}}' 18
        }
        Step "  placing the cover logo"
        $cover = PlaceLogo $doc.Content '{{LOGO_COVER}}' 30
        Step "  logo placed: sheets $($sheets -join ','), cover $cover"
        # The cover's corner pattern: on the cover page, behind the text, in the top-right corner, level with the logo.
        $pattern = [string](Resolve-Path (Join-Path $here 'out\cover-pattern.svg')).Path
        $size = [float]420   # points: about 5.8 in
        $pageWidth = [float]$section.PageSetup.PageWidth
        # Anchored to the cover's first paragraph (not in a table cell, where it would be placed relative to the cell).
        $anchor = $doc.Paragraphs.Item(1).Range
        $shape = $doc.Shapes.AddPicture($pattern, $false, $true, [float]0, [float]0, $size, $size, $anchor)
        $shape.LockAspectRatio = -1
        $shape.LayoutInCell = 0
        $shape.WrapFormat.Type = 5                # behind text
        $shape.RelativeHorizontalPosition = 1     # page
        $shape.RelativeVerticalPosition = 1       # page
        # No bleed: it must print on an office printer, so it stays inside the margins: right edge on the
        # right margin, top level with the top of the logo (top margin + the 0.25 in drop above the logo).
        $shape.Left = $pageWidth - [float]$section.PageSetup.RightMargin - $size
        $shape.Top = [float]$section.PageSetup.TopMargin + [float]18
        $shape.Name = 'Cover pattern'
        Step ("  cover pattern placed at {0:F0},{1:F0} ({2:F0} x {3:F0} pt)" -f $shape.Left, $shape.Top, $shape.Width, $shape.Height)
        if ($doc.TablesOfContents.Count -gt 0) { [void](Call $doc.TablesOfContents.Item(1) 'Update'); Step "  contents updated" }
        $window = $doc.ActiveWindow
        $window.View.Type = 3
        [void]$doc.ComputeStatistics(2)  # repaginate (wdStatisticPages), then refresh page fields in every story
        foreach ($story in @($doc.StoryRanges)) { $s = $story; while ($null -ne $s) { [void]$s.Fields.Update(); $s = $s.NextStoryRange } }
        foreach ($sec in @($doc.Sections)) { foreach ($k in 1, 2, 3) { [void]$sec.Headers.Item($k).Range.Fields.Update(); [void]$sec.Footers.Item($k).Range.Fields.Update() } }
        $pages = $window.ActivePane.Pages
        Step "  pages: $($pages.Count)"
        for ($n = 1; $n -le $pages.Count; $n++) { PagePicture $pages.Item($n) (Join-Path $here ("out\{0}-page{1}.png" -f $job.Pics, $n)) }
        $target = Join-Path $here ('out\' + $job.Out)
        [void](Call $doc 'SaveAs2' @([string]$target, [int]$job.Format))
        Step "  saved $target"
        [void](Call $doc 'Close' @(0))
    }
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
