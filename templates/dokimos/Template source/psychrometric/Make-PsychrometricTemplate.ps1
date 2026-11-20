<#
.SYNOPSIS
  Builds "Dokimos Psychrometric Chart v<N>.dotx" from the Dokimos design calculation template
  and the Calcpad script "Psychrometric chart.cpd".
.DESCRIPTION
  The design inputs become the site and the two known properties (in place of materials,
  geometry and loads); the calculations get the state of the air and the chart. See
  ..\TemplateBuild.ps1. Writes steps.log and pid.txt to OutDir, the template and page pictures
  to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 3)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$name = "Dokimos Psychrometric Chart v$Version"
try {
    if (-not (Open-Template $OutDir -Script (Join-Path $PSScriptRoot 'Psychrometric chart.cpd'))) { return }

    (Tagged 'dokimos:title').Range.Text = 'Psychrometric chart'
    (Tagged 'dokimos:codes').Range.Text = 'ASHRAE Handbook – Fundamentals, chapter 1'
    [void](FillSection 'Scope' @(
        'Properties of moist air from two known properties, at the standard atmospheric pressure of the site, with a psychrometric chart.'))
    [void](FillSection 'Codes and standards' @(
        'ASHRAE Handbook – Fundamentals (SI), chapter 1, Psychrometrics: standard atmosphere, saturation pressure (Hyland and Wexler), moist air relations.'))
    [void](FillSection 'Assumptions' @(
        'Moist air is an ideal-gas mixture of dry air and water vapour, at the standard atmospheric pressure for the altitude.',
        'Thermodynamic wet-bulb temperature; saturation over ice below 0 °C.',
        'Chart from 0 to 50 °C and up to 30 g/kg dry air, widened in steps of 5 to show the air.'))

    RenameHeading 'Materials' 'Site'
    RenameHeading 'Geometry' 'Known properties of the air'
    RemoveHeading 'Loads'
    ConvertSection 'Site' 'site'
    ConvertSection 'Known properties of the air' 'known'

    AddHeadingAfter 'Calculations' 'State of the air'
    ConvertSection 'State of the air' 'calculations'
    AddSectionBefore 'Results and conclusions' 'Psychrometric chart'
    ConvertSection 'Psychrometric chart' 'chart'
    RemoveHeading 'Results and conclusions' # the working gives every value; a table would repeat it

    $properties = @(
        @('Dry-bulb temperature, °C', '1'), @('Relative humidity, %', '2'), @('Wet-bulb temperature, °C', '3'),
        @('Dew-point temperature, °C', '4'), @('Humidity ratio, g/kg dry air', '5'),
        @('Specific enthalpy, kJ/kg dry air', '6'), @('Specific volume, m³/kg dry air', '7'))
    Dropdown 'P_1' $properties '1' 'Known property: '
    Dropdown 'P_2' $properties '2' 'Known property: '
    JoinInputs 'P_1' 'x_1' ' = '
    JoinInputs 'P_2' 'x_2' ' = '

    $out = Join-Path $PSScriptRoot 'out'
    Save-Template (Join-Path $out "$name.dotx") $out "v$Version"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
