<#
.SYNOPSIS
  Runs Tekla Tedds' "RC beam design (EN1992)" for the beam template's default section, without
  Tedds' window, and saves what Tedds wrote: output.txt (its calculation as text), output.rtf and
  output.xml (every variable it ended with). BeamTemplateTests' Tedds checks come from these runs.
.DESCRIPTION
  Uses Tedds' own API: the "Tedds.Calculator" COM server (TeddsCalc.exe, 32-bit, reached late-bound
  from 64-bit PowerShell). The recipe and variable names come from Tekton's Tedds integration
  (measured on Tedds 28.1): concrete by class (_ConcreteClass), section s1, φ written \66.
  A run that hits an error can leave a Tedds dialog waiting on screen. So run this script
  through a watchdog that stops only the TeddsCalc process the run started (listed in pid.txt),
  never a Tedds the user has open.
  Section: 300 × 500, C30/37, cover 30, 3 Ø20 bottom, 2 Ø12 top, Ø8 links (2 legs) at 200;
  M_Ed = 137.53 kNm, V_Ed = 91.69 kN, M_qp = 72.68 kNm (the template's default beam).
.PARAMETER Recommended
  Use the recommended values (α_cc = 1.0) instead of Tedds' default, the UK annex (α_cc = 0.85).
#>
param(
    [Parameter(Mandatory)][string]$OutDir,
    [switch]$Recommended
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$log = Join-Path $OutDir 'steps.log'; Remove-Item $log -ErrorAction SilentlyContinue
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$m, [object[]]$a = @()) { [System.__ComObject].InvokeMember($m, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $a) }
function Prop($obj, [string]$p) { [System.__ComObject].InvokeMember($p, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }

# ---- Variables, in Tedds' units (SI base units) ----------------------------------------
$vars = [ordered]@{}
function Dim([string]$n, [double]$v, [int]$mass = 0, [int]$length = 0, [int]$time = 0) {
    $a = ''; if ($mass) { $a += " Mass=`"$mass`"" }; if ($length) { $a += " Length=`"$length`"" }; if ($time) { $a += " Time=`"$time`"" }
    $vars[$n] = [pscustomobject]@{ Type = 'Dim'; Raw = $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture); Attrs = $a }
}
function Txt([string]$n, [string]$t) { $vars[$n] = [pscustomobject]@{ Type = 'String'; Raw = '"' + $t.Replace('"', '\"') + '"'; Attrs = '' } }
function Mm([string]$n, [double]$mm) { Dim $n ($mm / 1000) -length 1 }
function KNm([string]$n, [double]$v) { Dim $n ($v * 1000) -mass 1 -length 2 -time -2 }
function KN([string]$n, [double]$v) { Dim $n ($v * 1000) -mass 1 -length 1 -time -2 }

Dim '_CalcUI' 0 # no window
if ($Recommended) { Txt 'NationalAnnex' 'Recommended'; Txt '_NationalAnnexItem' 'NDP Recommended' }
Txt '_ConcreteClass' 'C30/37'; Txt '_ConcreteClassMain' 'C30/37'
Mm 'b' 300; Mm 'h' 500
Mm 'c_{nom_t}' 30; Mm 'c_{nom_b}' 30; Mm 'c_{nom_s}' 30
Txt 'Description_{s1}' 'Mid-span'; Txt '_SectionType_{s1}' 'Rectangular'
Mm 'b_{s1}' 300; Mm 'h_{s1}' 500
KNm 'M_{pos_s1}' 137.53; KNm 'M_{neg_s1}' 0
KNm 'M_{pos_QP_s1}' 72.68; KNm 'M_{neg_QP_s1}' 0
KN 'V_{s1}' 91.69; KN 'V_{Ed,max_s1}' 91.69
Dim 'N_{s1_bL}' 1; Dim 'N_{s1_add_bL}' 1; Dim 'N_{s1_b_L1}' 3; Mm '\66_{s1_b_L1}' 20; Dim 'N_{s1_add_b_L1}' 0; Mm '\66_{s1_add_b_L1}' 20
Dim 'N_{s1_tL}' 1; Dim 'N_{s1_add_tL}' 1; Dim 'N_{s1_t_L1}' 2; Mm '\66_{s1_t_L1}' 12; Dim 'N_{s1_add_t_L1}' 0; Mm '\66_{s1_add_t_L1}' 12
Dim 'N_{s1_v}' 2; Mm '\66_{s1_v}' 8; Mm 's_{s1_v}' 200

$sb = New-Object Text.StringBuilder
[void]$sb.Append('<VariableSections xmlns="http://www.cscworld.com/Tedds/VariableSections"><VariableSection Id="0"><Variables>')
foreach ($name in @($vars.Keys | Sort-Object { $_ } -CaseSensitive)) {
    $v = $vars[[string]$name]
    [void]$sb.Append("<Variable><Name>$([Security.SecurityElement]::Escape([string]$name))</Name><Value Type=`"$($v.Type)`"$($v.Attrs)>$([Security.SecurityElement]::Escape($v.Raw))</Value></Variable>")
}
[void]$sb.Append('</Variables></VariableSection></VariableSections>')
$xml = $sb.ToString()

# ---- Run --------------------------------------------------------------------------------
$before = @(Get-Process TeddsCalc -ErrorAction SilentlyContinue | ForEach-Object Id)
$calc = [Activator]::CreateInstance([Type]::GetTypeFromProgID('Tedds.Calculator'))
$mine = @(Get-Process TeddsCalc -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
Set-Content (Join-Path $OutDir 'pid.txt') ($mine -join ',')
Step ("calculator started: TeddsCalc {0}" -f ($mine -join ','))
try {
    [void](Call $calc 'InitializeCalc' @('$(SysLbrDir)RC member design-EN1992-si-engb.lbr', 'RC beam design', $xml, '', 0, 0, 0))
    Step "status: $(Prop $calc 'Status') (1 = calculated)"
    [IO.File]::WriteAllText((Join-Path $OutDir 'output.xml'), [string](Call $calc 'GetVariables' @(1)), (New-Object Text.UTF8Encoding($false)))
    $rtf = [string](Call $calc 'GetOutput' @(1))
    [IO.File]::WriteAllText((Join-Path $OutDir 'output.rtf'), $rtf, [Text.Encoding]::ASCII)
    Add-Type -AssemblyName System.Windows.Forms
    $box = New-Object System.Windows.Forms.RichTextBox
    $box.Rtf = $rtf
    [IO.File]::WriteAllText((Join-Path $OutDir 'output.txt'), $box.Text, (New-Object Text.UTF8Encoding($false)))
    Step 'saved output.txt, output.rtf, output.xml'
}
catch { Step ("FAILED: " + $_.Exception.InnerException.Message + ' / ' + $_.Exception.Message) }
finally {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($calc)
    Start-Sleep -Seconds 2
    foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force; Step "stopped TeddsCalc $id" } }
}
