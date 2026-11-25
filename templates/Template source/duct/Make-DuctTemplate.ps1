<#
.SYNOPSIS
  Builds "Duct Sizing v<N>.dotx" (and the Dokimos one, where its base is) from the design calculation template and the
  Calcpad script "Duct sizing.cpd". It replaces McQuay's DuctSizer.
.DESCRIPTION
  The design inputs become the air and the duct (in place of materials, geometry and loads); the
  calculations get the air properties, the round duct and the rectangular duct. See
  ..\TemplateBuild.ps1. Writes steps.log and pid.txt to OutDir, the template and page pictures to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 2)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$name = TemplateName "Duct Sizing v$Version"
try {
    if (-not (Open-Template $OutDir -Script (Join-Path $PSScriptRoot 'Duct sizing.cpd'))) { return }

    (Field 'title').Range.Text = 'Duct sizing'
    (Field 'codes').Range.Text = 'ASHRAE Handbook – Fundamentals, chapter 21'
    [void](FillSection 'Scope' @(
        'Round and rectangular air duct for a given airflow, sized by friction rate, velocity or diameter.'))
    [void](FillSection 'Codes and standards' @(
        'ASHRAE Handbook – Fundamentals (SI), chapter 21, Duct design: Darcy–Weisbach and Colebrook, duct roughness, equivalent diameter of rectangular ducts (Huebscher); chapter 1: standard atmosphere.'))
    [void](FillSection 'Assumptions' @(
        'Dry air at the standard atmospheric pressure of the site; humidity changes the density by less than 1 %.',
        'Straight duct with fully developed turbulent flow; fittings are not included.'))

    RenameHeading 'Materials' 'Air'
    RenameHeading 'Geometry' 'Duct'
    RemoveHeading 'Loads'
    ConvertSection 'Air' 'air'
    ConvertSection 'Duct' 'duct'

    AddHeadingAfter 'Calculations' 'Air properties'
    ConvertSection 'Air properties' 'properties'
    AddSectionBefore 'Results and conclusions' 'Round duct'
    ConvertSection 'Round duct' 'round'
    AddSectionBefore 'Results and conclusions' 'Rectangular duct'
    ConvertSection 'Rectangular duct' 'rectangular'
    RemoveHeading 'Results and conclusions' # the working gives the sizes; a summary would repeat them

    # ASHRAE's roughness categories (Fundamentals, chapter 21, Table 1).
    Dropdown 'ε' @(
        @('Smooth, e.g. PVC or aluminium (0.03 mm)', '0.03'),
        @('Medium smooth, e.g. galvanized steel (0.09 mm)', '0.09'),
        @('Average (0.15 mm)', '0.15'),
        @('Medium rough, e.g. rigid fibrous glass duct (0.9 mm)', '0.9'),
        @('Rough, e.g. flexible duct or concrete (3 mm)', '3')) '0.09' 'Duct material: '
    Dropdown 'P_s' @(@('Friction rate, Pa/m', '1'), @('Velocity, m/s', '2'), @('Round duct diameter, mm', '3')) '1' 'Size by: '
    JoinInputs 'P_s' 'x_s' ' = '

    $out = Join-Path $PSScriptRoot 'out'
    Save-Template (Join-Path $out "$name.dotx") $out "v$Version"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
