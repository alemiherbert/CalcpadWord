<#
.SYNOPSIS
  Builds "Dokimos RC Beam Design (EC2) v<N>.dotx" from the Dokimos design calculation template
  and the Calcpad script "RC beam design (EN 1992-1-1).cpd".
.DESCRIPTION
  Each section of the script goes under its heading and is converted by
  Calcpad for Word: text becomes Word paragraphs, calculations inline equations, checks small
  Calcpad blocks, and every "name = ? {default}" a blue input field. Some inputs then become
  drop-down lists. See ..\TemplateBuild.ps1. Writes steps.log and pid.txt to OutDir, the template
  and page pictures to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 3)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$name = "Dokimos RC Beam Design (EC2) v$Version"
try {
    if (-not (Open-Template $OutDir -Script (Join-Path $PSScriptRoot 'RC beam design (EN 1992-1-1).cpd'))) { return }

    (Tagged 'dokimos:title').Range.Text = 'Reinforced concrete beam design'
    (Tagged 'dokimos:codes').Range.Text = 'EN 1990, EN 1991-1-1, EN 1992-1-1, EN 1992-1-2'
    [void](FillSection 'Scope' @(
        'Simply supported rectangular RC beam under uniform load, to EN 1992-1-1: cover, fire, bending, shear, cracking and deflection.'))
    [void](FillSection 'Codes and standards' @(
        'EN 1990:2002 + A1:2005: combinations (6.10) and (6.16).',
        'EN 1991-1-1:2002: self-weight and imposed loads.',
        'EN 1992-1-1:2004 + A1:2014: recommended values unless stated.',
        'EN 1992-1-2:2004: tabulated data for beams (5.6).'))
    [void](FillSection 'Assumptions' @(
        'Linear elastic analysis without redistribution; shear checked at the support.',
        'Concrete C50/60 or lower; bottom bars in one layer, continued to the supports; vertical links.',
        'Cover for structural class S4, without reductions. Fire on three sides.',
        'Creep and shrinkage for cement class N, loaded at 28 days, the whole perimeter drying.'))

    ConvertSection 'Materials' 'materials'
    ConvertSection 'Geometry' 'geometry'
    ConvertSection 'Loads' 'loads'
    ConvertSection 'Calculations' 'calculations'
    ConvertSection 'Results and conclusions' 'results'

    Dropdown 'f_ck' @(@('C20/25', '20'), @('C25/30', '25'), @('C30/37', '30'), @('C35/45', '35'), @('C40/50', '40'), @('C45/55', '45'), @('C50/60', '50')) '30' 'Strength class: '
    # The exposure class gives the minimum cover for durability (S4, Table 4.4N); w_max follows from it.
    Dropdown 'c_min,dur' @(@('X0', '10'), @('XC1', '15'), @('XC2, XC3', '25'), @('XC4', '30'), @('XD1, XS1', '35'), @('XD2, XS2', '40'), @('XD3, XS3', '45')) '15' 'Exposure class: '
    Dropdown 'RH' @(@('Inside buildings, RH 50%', '50'), @('Outside, RH 80%', '80')) '50' 'Environment: '
    Dropdown 'R' @(@('R30', '30'), @('R60', '60'), @('R90', '90'), @('R120', '120'), @('R180', '180'), @('R240', '240')) '60' 'Fire resistance: '
    Dropdown 'P_b' @(@('No', '0'), @('Yes', '1')) '0' 'Partitions liable to damage: '
    Dropdown 'ψ_2' @(@('A: domestic and residential', '0.3'), @('B: offices', '0.30'), @('C: congregation areas', '0.6'), @('D: shopping areas', '0.60'), @('E: storage areas', '0.8'), @('H: roofs', '0')) '0.3' 'Category of use (EN 1991-1-1): '
    $bars = @('8', '10', '12', '16', '20', '25', '32') | ForEach-Object { , @($_, $_) }
    Dropdown 'Φ_t' $bars '20' $null
    Dropdown 'Φ_c' $bars '12' $null
    Dropdown 'Φ_w' (@('6', '8', '10', '12') | ForEach-Object { , @($_, $_) }) '8' $null

    $out = Join-Path $PSScriptRoot 'out'
    Save-Template (Join-Path $out "$name.dotx") $out "v$Version"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
