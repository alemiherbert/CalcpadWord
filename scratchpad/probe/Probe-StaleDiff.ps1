param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
function Idle() { for ($w = 0; $w -lt 600 -and (Prop $api 'IsWorking'); $w++) { Start-Sleep -Milliseconds 100 } }
function Body($doc) { [regex]::Replace(([regex]::Match($doc.Content.WordOpenXML, '<w:body>.*</w:body>', 'Singleline').Value), ' w:rsid\w*="[^"]*"|<w:docPart w:val="[^"]*"/>', '') }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    Idle
    $fixture = 'c:\Users\dell alienware\Desktop\Alemi Herbert\Plugin\tests\CalcpadWord.Engine.Tests\Fixtures\SoilStress.cpd'
    $soil = @([IO.File]::ReadAllText($fixture, [Text.Encoding]::UTF8).TrimEnd() -split "`r?`n")
    $doc = $word.Documents.Add()
    $doc.Content.Text = ($soil -join "`r")
    $doc.Content.Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate')); Idle; Step (Prop $api 'LastStatus')
    $a = Body $doc
    [IO.File]::WriteAllText((Join-Path $OutDir 'a.xml'), $a)
    $settings = Call $api 'GetSettings'
    Step ('SetSettings 4 -> ' + (Call $api 'SetSettings' @($settings -replace 'decimals=\d+', 'decimals=4')))
    Step ('Cancel -> ' + (Call $api 'Cancel')); Idle
    Step ('SetSettings back -> ' + (Call $api 'SetSettings' @($settings))); Idle; Step (Prop $api 'LastStatus')
    $b = Body $doc
    [IO.File]::WriteAllText((Join-Path $OutDir 'b.xml'), $b)
    Step ("identical: {0}; lengths {1} / {2}" -f ($a -eq $b), $a.Length, $b.Length)
    if ($a -ne $b) {
        $k = 0; while ($k -lt [Math]::Min($a.Length, $b.Length) -and $a[$k] -eq $b[$k]) { $k++ }
        $from = [Math]::Max(0, $k - 300)
        Step ("first difference at $k`n  was: " + $a.Substring($from, [Math]::Min(700, $a.Length - $from)) + "`n  now: " + $b.Substring($from, [Math]::Min(700, $b.Length - $from)))
    }
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
