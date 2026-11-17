<#
.SYNOPSIS
  Builds "Dokimos RC Beam Design (EC2) v<N>.dotx" from the Dokimos design calculation template
  and the Calcpad script "RC beam design (EN 1992-1-1).cpd".
.DESCRIPTION
  Runs in its own hidden Word (your open documents are left alone) with the Calcpad for Word
  add-in (0.8.0 or later: input fields). Each section of the script (split at its
  '<!--section:NAME--> lines) replaces the guidance under the matching heading and is converted
  by the add-in's Evaluate: text becomes Word paragraphs, calculations inline equations, checks
  small Calcpad blocks, and every "name = ? {default}" a blue input field. Some inputs then
  become drop-down lists (concrete class, exposure class, category of use, bar sizes).
  Writes steps.log and pid.txt (the hidden Word's process ID) to OutDir, page pictures to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 1)
$ErrorActionPreference = 'Stop'
$here = Split-Path $PSCommandPath
$base = (Resolve-Path (Join-Path $here '..\..\Dokimos Design Calculation v3.dotx')).Path
$script = Join-Path $here 'RC beam design (EN 1992-1-1).cpd'
$name = "Dokimos RC Beam Design (EC2) v$Version"
New-Item -ItemType Directory -Force $OutDir, (Join-Path $here 'out') | Out-Null
$log = Join-Path $OutDir 'steps.log'; Remove-Item $log -ErrorAction SilentlyContinue
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$method, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$property) { [System.__ComObject].InvokeMember($property, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
# The first content control with this tag.
function Tagged([string]$tag) {
    $all = $doc.ContentControls
    for ($i = 1; $i -le $all.Count; $i++) { $c = $all.Item($i); if ($c.Tag -eq $tag) { return $c } }
    throw "No content control tagged $tag"
}
Add-Type -AssemblyName System.Drawing

# The script's sections.
$sections = @{}; $current = $null
foreach ($line in [IO.File]::ReadAllLines($script, [Text.Encoding]::UTF8)) {
    if ($line -match "^'<!--section:(\w+)-->\s*$") { $current = $Matches[1]; $sections[$current] = New-Object Collections.Generic.List[string]; continue }
    if ($current) { $sections[$current].Add($line) }
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate Word instance: stopping'; return }
$wordPid = $new[0]
Set-Content (Join-Path $OutDir 'pid.txt') $wordPid

function WaitForCalcpad([string]$what) {
    for ($i = 0; $i -lt 480 -and (Prop $api 'IsWorking'); $i++) { Start-Sleep -Milliseconds 250 }
    Step ("{0}: {1}" -f $what, (Prop $api 'LastStatus'))
}

function HeadingParagraph([string]$text) {
    $paragraphs = $doc.Paragraphs
    for ($i = 1; $i -le $paragraphs.Count; $i++) {
        $p = $paragraphs.Item($i)
        if ($p.OutlineLevel -le 2 -and $p.Range.Text.Trim() -eq $text) { return $p }
    }
    throw "Heading '$text' not found"
}

# Replaces the guidance paragraph after a heading with paragraphs of text (in the guidance's style); returns their range.
function ReplaceGuidance([string]$heading, [string[]]$lines) {
    $guide = (HeadingParagraph $heading).Next(1)
    $controls = $guide.Range.ContentControls
    for ($k = $controls.Count; $k -ge 1; $k--) { $controls.Item($k).Delete($true) }
    $range = $guide.Range
    [void]$range.MoveEnd(1, -1) # not the paragraph mark
    $range.Text = ($lines -join "`r")
    return $range
}

function ConvertSection([string]$heading, [string]$section) {
    $range = ReplaceGuidance $heading $sections[$section].ToArray()
    $range.Select()
    $status = Call $api 'Evaluate'
    WaitForCalcpad "$heading ($section): $status"
}

# Turns an input field into a drop-down list; the text before it in its paragraph becomes $label.
function Dropdown([string]$variable, [object[]]$entries, [string]$default, [string]$label) {
    $cc = Tagged "calcpad-input:$variable"
    if ($label) {
        $paragraph = $cc.Range.Paragraphs.Item(1).Range
        $doc.Range($paragraph.Start, $cc.Range.Start - 1).Text = $label
    }
    $cc.Type = 4 # wdContentControlDropdownList
    foreach ($e in $entries) { [void]$cc.DropdownListEntries.Add($e[0], $e[1]) }
    $list = $cc.DropdownListEntries
    for ($k = 1; $k -le $list.Count; $k++) { if ($list.Item($k).Value -eq $default) { $list.Item($k).Select(); break } }
    Step ("drop-down {0}: {1} entries, shows '{2}'" -f $variable, $list.Count, $cc.Range.Text)
}

try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    Step "Calcpad for Word $(Prop $api 'Version')"
    $doc = Call $word.Documents 'Add' @($base)

    # Cover
    (Tagged 'dokimos:title').Range.Text = 'Reinforced concrete beam design'
    (Tagged 'dokimos:codes').Range.Text = 'EN 1990, EN 1991-1-1, EN 1992-1-1'

    # Scope, design basis
    [void](ReplaceGuidance 'Scope' @(
        'Design of a simply supported reinforced concrete beam of rectangular section, uniformly loaded, to EN 1992-1-1: bending, shear, reinforcement limits and bar spacing at the ultimate limit state; crack width and deflection at the serviceability limit state.',
        'To use it, fill in the blue fields under Design inputs: type a value, or pick one from the list. The calculation updates when you leave a field, and the results and conclusions show whether the beam passes.'))
    [void](ReplaceGuidance 'Codes and standards' @(
        'EN 1990:2002 + A1:2005, Basis of structural design: load combinations (6.10) and (6.16).',
        'EN 1991-1-1:2002, Actions on structures: densities, self-weight and imposed loads.',
        'EN 1992-1-1:2004 + A1:2014, Design of concrete structures: recommended values of the nationally determined parameters, except where stated.'))
    [void](ReplaceGuidance 'Assumptions' @(
        'The beam is simply supported over the effective span, with a uniformly distributed load; the analysis is linear elastic, without redistribution.',
        'The bottom bars are in one layer and continue to the supports; the links are vertical, of the same steel as the main bars.',
        'Self-weight is from a density of 25 kN/m³. Creep is taken as the final creep coefficient given under Materials.',
        'Deflection is checked by the span to effective depth ratio (7.4.2); calculate it where the limit is approached or partitions are sensitive.'))

    # Design inputs and calculations, from the Calcpad script
    ConvertSection 'Materials' 'materials'
    ConvertSection 'Geometry' 'geometry'
    ConvertSection 'Loads' 'loads'
    ConvertSection 'Calculations' 'calculations'
    ConvertSection 'Results and conclusions' 'results'

    # The guidance's style (a shaded box) stays on the empty paragraph after the results: make it plain.
    $paragraphs = $doc.Paragraphs
    for ($i = $paragraphs.Count; $i -ge 1; $i--) {
        $p = $paragraphs.Item($i)
        if ($p.Style.NameLocal -match '^Result ?summary$' -and $p.Range.Text.Trim().Length -eq 0) { $p.Style = -1; Step 'cleared the empty result summary paragraph' }
    }

    # Inputs that are choices
    Dropdown 'f_ck' @(@('C20/25', '20'), @('C25/30', '25'), @('C30/37', '30'), @('C35/45', '35'), @('C40/50', '40'), @('C45/55', '45'), @('C50/60', '50')) '30' 'Strength class: '
    Dropdown 'w_max' @(@('X0, XC1', '0.4'), @('XC2, XC3, XC4', '0.3'), @('XD1, XD2, XS1, XS2, XS3', '0.30')) '0.3' 'Exposure class: '
    Dropdown 'ψ_2' @(@('A: domestic and residential', '0.3'), @('B: offices', '0.30'), @('C: congregation areas', '0.6'), @('D: shopping areas', '0.60'), @('E: storage areas', '0.8'), @('H: roofs', '0')) '0.3' 'Category of use (EN 1991-1-1): '
    $bars = @('8', '10', '12', '16', '20', '25', '32') | ForEach-Object { , @($_, $_) }
    Dropdown 'Φ_t' $bars '20' $null
    Dropdown 'Φ_c' $bars '12' $null
    Dropdown 'Φ_w' (@('6', '8', '10', '12') | ForEach-Object { , @($_, $_) }) '8' $null

    $status = Call $api 'Recalculate'
    WaitForCalcpad "recalculated: $status"
    Step ("errors: {0}" -f (Prop $api 'ErrorCount'))

    if ($doc.TablesOfContents.Count -gt 0) { [void](Call $doc.TablesOfContents.Item(1) 'Update') }
    $window = $doc.ActiveWindow; $window.View.Type = 3
    [void]$doc.ComputeStatistics(2)
    foreach ($sec in @($doc.Sections)) { [void]$sec.Headers.Item(1).Range.Fields.Update() }
    $pages = $window.ActivePane.Pages
    Step "pages: $($pages.Count)"
    for ($n = 1; $n -le $pages.Count; $n++) {
        $emf = Join-Path $here ("out\v{0}-page{1}.emf" -f $Version, $n); [IO.File]::WriteAllBytes($emf, [byte[]]$pages.Item($n).EnhMetaFileBits)
        $mf = New-Object System.Drawing.Imaging.Metafile($emf); $bmp = New-Object System.Drawing.Bitmap(1200, [int]($mf.Height * 1200.0 / $mf.Width))
        $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); $g.SmoothingMode = 'HighQuality'; $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
        $bmp.Save([IO.Path]::ChangeExtension($emf, '.png'), [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
    }
    $target = [string](Join-Path $here "out\$name.dotx")
    [void](Call $doc 'SaveAs2' @($target, 14))
    Step "saved $target"
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally {
    try { [void](Call $word 'Quit' @(0)) } catch { }
    Start-Sleep -Seconds 3
    if (Get-Process -Id $wordPid -ErrorAction SilentlyContinue) { Stop-Process -Id $wordPid -Force; Step "stopped hidden Word $wordPid" }
}
