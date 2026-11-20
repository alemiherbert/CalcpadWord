<#
.SYNOPSIS
  Shared by the scripts that build and test calculation templates (beam\, psychrometric\).
.DESCRIPTION
  A calculation template is the Dokimos design calculation template (v4) with a Calcpad script
  in it. Dot-source this file, then, inside try / finally:
    Open-Template $OutDir $base $script   # a hidden Word of its own; a new document from $base;
                                          # the script's sections ('<!--section:NAME--> lines)
    ...  the template's own steps: Tagged, FillSection, ConvertSection, Dropdown, ...
    Save-Template $target $pictures $prefix   # calculate, update the contents,
                                              # page pictures, save
    Close-Template                        # in finally: quits that Word, and stops only it if need be
  Your open documents are left alone. Each step is logged to OutDir\steps.log, and OutDir\pid.txt
  holds the hidden Word's process ID. It needs the Calcpad for Word add-in, 0.1.1 or later.
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$script:BaseTemplate = (Resolve-Path (Join-Path $PSScriptRoot '..\Dokimos Design Calculation v4.dotx')).Path

function Step([string]$m) { Add-Content -Path $script:log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$method, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$property) { [System.__ComObject].InvokeMember($property, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }

<# Starts a hidden Word of its own and makes a new document from $Template. False if Word gave an existing instance (then nothing is touched). #>
function Open-Template([string]$OutDir, [string]$Template = $script:BaseTemplate, [string]$Script = $null) {
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $script:log = Join-Path $OutDir 'steps.log'; Remove-Item $script:log -ErrorAction SilentlyContinue
    $script:sections = @{}
    if ($Script) {
        $current = $null
        foreach ($line in [IO.File]::ReadAllLines($Script, [Text.Encoding]::UTF8)) {
            if ($line -match "^'<!--section:(\w+)-->\s*$") { $current = $Matches[1]; $script:sections[$current] = New-Object Collections.Generic.List[string]; continue }
            if ($current) { $script:sections[$current].Add($line) }
        }
    }
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $script:word = New-Object -ComObject Word.Application
    $new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
    if ($new.Count -ne 1) { $script:word = $null; Step 'not a separate Word instance: stopping'; return $false }
    $script:wordPid = $new[0]
    Set-Content (Join-Path $OutDir 'pid.txt') $script:wordPid
    $script:word.Visible = $false
    $script:word.DisplayAlerts = 0
    $script:api = $script:word.COMAddIns.Item('CalcpadWord.Connect').Object
    Step "Calcpad for Word $(Prop $script:api 'Version')"
    $script:doc = Call $script:word.Documents 'Add' @($Template)
    Step "new document from $(Split-Path $Template -Leaf)"
    return $true
}

function Close-Template {
    if (-not $script:wordPid) { return }
    try { [void](Call $script:word 'Quit' @(0)) } catch { }
    Start-Sleep -Seconds 3
    if (Get-Process -Id $script:wordPid -ErrorAction SilentlyContinue) { Stop-Process -Id $script:wordPid -Force; Step "stopped hidden Word $script:wordPid" }
}

<# The first content control with this tag. #>
function Tagged([string]$tag) {
    $all = $script:doc.ContentControls
    for ($i = 1; $i -le $all.Count; $i++) { $c = $all.Item($i); if ($c.Tag -eq $tag) { return $c } }
    throw "No content control tagged $tag"
}

function WaitForCalcpad([string]$what) {
    for ($i = 0; $i -lt 480 -and (Prop $script:api 'IsWorking'); $i++) { Start-Sleep -Milliseconds 250 }
    $status = Prop $script:api 'LastStatus'
    if ($what) { Step ("{0}: {1}" -f $what, $status) }
    return $status
}

<# Waits until Calcpad is idle and $condition holds (true), or $seconds pass (false). Word runs deferred work, such as calculating after a field is left, while it waits. #>
function WaitUntil([scriptblock]$condition, [int]$seconds = 60) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $seconds) {
        if (-not (Prop $script:api 'IsWorking') -and (& $condition)) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

<# True when every Calcpad expression and block is locked and no input field is. #>
function CalculationLocked {
    $all = $script:doc.ContentControls
    for ($i = 1; $i -le $all.Count; $i++) {
        $c = $all.Item($i)
        if ($c.Tag -like 'calcpad-input:*') { if ($c.LockContentControl) { return $false } }
        elseif ($c.Tag -like 'calcpad*' -and -not $c.LockContentControl) { return $false }
    }
    return $true
}

function HeadingParagraph([string]$text) {
    $paragraphs = $script:doc.Paragraphs
    for ($i = 1; $i -le $paragraphs.Count; $i++) {
        $p = $paragraphs.Item($i)
        if ($p.OutlineLevel -le 2 -and $p.Range.Text.Trim() -eq $text) { return $p }
    }
    throw "Heading '$text' not found"
}

function SetText($paragraph, [string]$text) {
    $range = $paragraph.Range
    [void]$range.MoveEnd(1, -1) # not the paragraph mark
    $range.Text = $text
}

<# Fills the empty paragraph after a heading with paragraphs of text (in its style: bullets under Codes and Assumptions); returns their range. #>
function FillSection([string]$heading, [string[]]$lines) {
    $body = (HeadingParagraph $heading).Next(1)
    $range = $body.Range
    [void]$range.MoveEnd(1, -1)
    $range.Text = ($lines -join "`r")
    return $range
}

<# Puts a section of the script under a heading and lets Calcpad for Word convert it. #>
function ConvertSection([string]$heading, [string]$section) {
    $range = FillSection $heading $script:sections[$section].ToArray()
    $range.Select()
    $status = Call $script:api 'Evaluate'
    [void](WaitForCalcpad "$heading ($section): $status")
}

function RenameHeading([string]$old, [string]$new) { SetText (HeadingParagraph $old) $new }

<# Removes a heading and the empty paragraph under it. #>
function RemoveHeading([string]$text) {
    $heading = HeadingParagraph $text
    $body = $heading.Next(1)
    $script:doc.Range($heading.Range.Start, $body.Range.End).Delete() | Out-Null
}

<# A level-2 heading straight after the heading $after (before its empty paragraph, which then belongs to the new one). #>
function AddHeadingAfter([string]$after, [string]$text) {
    $parent = HeadingParagraph $after
    $parent.Range.InsertParagraphAfter()
    $new = $parent.Next(1)
    $new.Style = -3 # wdStyleHeading2
    SetText $new $text
}

<# A level-2 heading and an empty paragraph for its section, just before the heading $before. #>
function AddSectionBefore([string]$before, [string]$text) {
    $next = HeadingParagraph $before
    $next.Range.InsertParagraphBefore()
    $next.Range.InsertParagraphBefore()
    $body = $next.Previous(1)
    $heading = $next.Previous(2)
    $heading.Style = -3 # wdStyleHeading2
    $body.Style = -1    # wdStyleNormal
    SetText $heading $text
}

<# Turns an input field into a drop-down list; the text before it in its paragraph becomes $label. #>
function Dropdown([string]$variable, [object[]]$entries, [string]$default, [string]$label) {
    $cc = Tagged "calcpad-input:$variable"
    if ($label) {
        $paragraph = $cc.Range.Paragraphs.Item(1).Range
        $script:doc.Range($paragraph.Start, $cc.Range.Start - 1).Text = $label
    }
    $cc.Type = 4 # wdContentControlDropdownList
    foreach ($e in $entries) { [void]$cc.DropdownListEntries.Add($e[0], $e[1]) }
    $list = $cc.DropdownListEntries
    for ($k = 1; $k -le $list.Count; $k++) { if ($list.Item($k).Value -eq $default) { $list.Item($k).Select(); break } }
    Step ("drop-down {0}: {1} entries, shows '{2}'" -f $variable, $list.Count, $cc.Range.Text)
}

<# Replaces the text between two input fields in one paragraph, e.g. "[property] = [value]". #>
function JoinInputs([string]$first, [string]$second, [string]$text) {
    $a = Tagged "calcpad-input:$first"
    $b = Tagged "calcpad-input:$second"
    $script:doc.Range($a.Range.End + 1, $b.Range.Start - 1).Text = $text
}

<# A picture of one page, 1200 pixels wide. #>
function Save-PagePicture([int]$page, [string]$path) {
    $window = $script:doc.ActiveWindow; $window.View.Type = 3
    $emf = [IO.Path]::ChangeExtension($path, '.emf')
    [IO.File]::WriteAllBytes($emf, [byte[]]$window.ActivePane.Pages.Item($page).EnhMetaFileBits)
    $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1200, [int]($mf.Height * 1200.0 / $mf.Width))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.SmoothingMode = 'HighQuality'; $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
}

<#
  Finishes a template: calculates, updates the contents and the sheet numbers, draws each page to
  $pictures\$prefix-pageN.png (replacing older ones) and saves it as $target (.dotx).
#>
function Save-Template([string]$target, [string]$pictures, [string]$prefix) {
    $status = Call $script:api 'Recalculate'
    [void](WaitForCalcpad "calculated: $status")
    # A section left for the designer (References) is an empty paragraph: made plain, so that no
    # lone bullet prints if it stays empty.
    $previous = $script:doc.Paragraphs.First
    $p = $previous.Next(1)
    while ($null -ne $p) {
        if ($previous.OutlineLevel -le 2 -and $p.Range.ListFormat.ListType -ne 0 -and $p.Range.Text.Trim().Length -eq 0) {
            $p.Style = -1 # wdStyleNormal
            Step ("empty section after '{0}' made plain" -f $previous.Range.Text.Trim())
        }
        $previous = $p
        $p = $p.Next(1)
    }
    # The empty paragraphs after the last block (one kept the summary's shaded style; Word's last
    # can't go) are made plain and 1 pt, so that they never add a blank page.
    $paragraphs = $script:doc.Paragraphs
    for ($i = $paragraphs.Count; $i -ge 1; $i--) {
        $p = $paragraphs.Item($i)
        if ($p.Range.Text.Trim().Length -gt 0 -or $p.Range.Information(12) -or $null -ne $p.Range.ParentContentControl) { break } # wdWithInTable
        # Only its mark: the paragraph just after a block also holds the block's (locked) end.
        $mark = $script:doc.Range($p.Range.End - 1, $p.Range.End)
        $mark.Style = -1 # wdStyleNormal
        $mark.Font.Size = 1
        $format = $mark.ParagraphFormat
        $format.SpaceBefore = 0; $format.SpaceAfter = 0
        $format.LineSpacingRule = 4; $format.LineSpacing = 1 # wdLineSpaceExactly, 1 pt
        Step "paragraph $i made 1 pt"
    }
    Step ("errors: {0}" -f (Prop $script:api 'ErrorCount'))
    # Calculate automatically stays on, the default (the user, 2026-11-21).
    Step ("settings: {0}" -f (Call $script:api 'GetSettings'))
    # The calculation is locked; only the input fields are filled in (the user, 2026-11-21).
    $script:doc.Content.Select()
    Step ("locked: {0}" -f (Call $script:api 'Lock'))
    Step ("calculation locked, inputs open: {0}" -f (CalculationLocked))

    if ($script:doc.TablesOfContents.Count -gt 0) { [void](Call $script:doc.TablesOfContents.Item(1) 'Update') }
    $window = $script:doc.ActiveWindow; $window.View.Type = 3
    [void]$script:doc.ComputeStatistics(2)
    foreach ($section in @($script:doc.Sections)) { [void]$section.Headers.Item(1).Range.Fields.Update() }
    New-Item -ItemType Directory -Force $pictures | Out-Null
    Get-ChildItem $pictures -Filter "$prefix-page*.png" | Remove-Item
    $count = $window.ActivePane.Pages.Count
    Step "pages: $count"
    for ($n = 1; $n -le $count; $n++) { Save-PagePicture $n (Join-Path $pictures "$prefix-page$n.png") }
    [void](Call $script:doc 'SaveAs2' @($target, 14))
    Step "saved $target"
    [void](Call $script:doc 'Close' @(0))
}
