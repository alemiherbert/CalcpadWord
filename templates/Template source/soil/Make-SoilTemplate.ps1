<#
.SYNOPSIS
  Builds "Soil Stresses and Settlement v<N>.dotx" (and the Dokimos one, where its base is) from the design calculation
  template and the Calcpad script "Soil stresses and settlement.cpd".
.DESCRIPTION
  The design inputs become the soil, the loaded area and the load; the calculations get the
  stresses, the settlement and the stresses in the plane through the centre. See
  ..\TemplateBuild.ps1. Writes steps.log and pid.txt to OutDir, the template and page pictures to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 2)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$name = TemplateName "Soil Stresses and Settlement v$Version"
try {
    if (-not (Open-Template $OutDir -Script (Join-Path $PSScriptRoot 'Soil stresses and settlement.cpd'))) { return }

    (Field 'title').Range.Text = 'Soil stresses and settlement'
    (Field 'codes').Range.Text = 'EN 1997-1, Annex F'
    [void](FillSection 'Scope' @(
        'Stresses and settlement under a rectangular surface load that varies linearly across its width.'))
    [void](FillSection 'Codes and standards' @(
        'EN 1997-1:2004, Annex F.1: settlement by the stress–strain method.',
        'Boussinesq: stresses under a point load on an elastic half-space, integrated over the loaded area.'))
    [void](FillSection 'Assumptions' @(
        'Homogeneous, isotropic, linear elastic soil (a half-space); flexible load on the surface.',
        'Settlement from the vertical strain (1 − ν²)·σ/E₀ of the vertical stress σ alone, down to the depth H.'))

    RenameHeading 'Materials' 'Soil'
    RenameHeading 'Geometry' 'Loaded area'
    ConvertSection 'Soil' 'soil'
    ConvertSection 'Loaded area' 'area'
    ConvertSection 'Loads' 'loads'
    # Each sketch stays on the page of the inputs it shows (an input with a unit is tagged with it).
    foreach ($tag in 'calcpad-input:H:m', 'calcpad-input:p_1:kPa') { (Tagged $tag).Range.Paragraphs.Item(1).KeepWithNext = -1 }

    AddHeadingAfter 'Calculations' 'Stresses (Boussinesq)'
    ConvertSection 'Stresses (Boussinesq)' 'stresses'
    AddSectionBefore 'Results and conclusions' 'Settlement'
    ConvertSection 'Settlement' 'settlement'
    AddSectionBefore 'Results and conclusions' 'Through the centre, across the width'
    ConvertSection 'Through the centre, across the width' 'plane'
    RemoveHeading 'Results and conclusions' # the working gives the settlements; a summary would repeat them

    $out = Join-Path $PSScriptRoot 'out'
    Save-Template (Join-Path $out "$name.dotx") $out "v$Version"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
