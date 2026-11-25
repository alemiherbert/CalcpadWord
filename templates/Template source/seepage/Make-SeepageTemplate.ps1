<#
.SYNOPSIS
  Builds "Seepage by Finite Differences v<N>.dotx" (and the Dokimos one, where its base is) from the design calculation
  template and the Calcpad script "Seepage by finite differences.cpd".
.DESCRIPTION
  The design inputs become the soil, the geometry and the water (with a sketch of the section);
  the calculations get the grid and its solution, the flow net, the exit gradient and heave, and
  the water pressure on the wall. See ..\TemplateBuild.ps1. Writes steps.log and pid.txt to OutDir,
  the template and page pictures to out\.
#>
param([Parameter(Mandatory)][string]$OutDir, [int]$Version = 1)
. (Join-Path $PSScriptRoot '..\TemplateBuild.ps1')
$name = TemplateName "Seepage by Finite Differences v$Version"
try {
    if (-not (Open-Template $OutDir -Script (Join-Path $PSScriptRoot 'Seepage by finite differences.cpd'))) { return }

    (Field 'title').Range.Text = 'Seepage under a sheet pile wall'
    (Field 'codes').Range.Text = 'EN 1997-1'
    [void](FillSection 'Scope' @(
        'Steady seepage under a sheet pile wall: heads, flow, exit gradient, heave and the water pressure on the wall.'))
    [void](FillSection 'Codes and standards' @(
        'EN 1997-1:2004, 2.4.7.5 and 10.3: failure by hydraulic heave (HYD), with the partial factors of Table A.17 (recommended values).',
        "Laplace’s equation for steady flow, solved by finite differences on a square grid."))
    [void](FillSection 'Assumptions' @(
        "Confined, steady flow in a vertical section (plane flow); Darcy’s law; saturated soil.",
        'Up to two homogeneous layers, each with the same ratio of horizontal to vertical permeability.',
        'Impermeable base, wall of no thickness; no flow across the ends of the section.'))

    RenameHeading 'Materials' 'Soil'
    RenameHeading 'Loads' 'Water'
    ConvertSection 'Soil' 'soil'
    ConvertSection 'Geometry' 'geometry'
    ConvertSection 'Water' 'water'
    # The sketch stays on the page of the water levels it shows (an input with a unit is tagged with it).
    (Tagged 'calcpad-input:h_1:m').Range.Paragraphs.Item(1).KeepWithNext = -1

    AddHeadingAfter 'Calculations' 'Finite differences'
    ConvertSection 'Finite differences' 'grid'
    AddSectionBefore 'Results and conclusions' 'Flow net'
    ConvertSection 'Flow net' 'flow'
    AddSectionBefore 'Results and conclusions' 'Exit gradient and heave'
    ConvertSection 'Exit gradient and heave' 'exit'
    AddSectionBefore 'Results and conclusions' 'Water pressure on the wall'
    ConvertSection 'Water pressure on the wall' 'wall'
    RemoveHeading 'Results and conclusions' # the checks give their verdicts; a summary would repeat them

    $out = Join-Path $PSScriptRoot 'out'
    Save-Template (Join-Path $out "$name.dotx") $out "v$Version"
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally { Close-Template }
