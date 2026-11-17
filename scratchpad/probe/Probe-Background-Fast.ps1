param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Prop($obj, [string]$name) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null) }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $fixture = 'c:\Users\dell alienware\Desktop\Alemi Herbert\Plugin\tests\CalcpadWord.Engine.Tests\Fixtures\SoilStress.cpd'
    $soil = @([IO.File]::ReadAllText($fixture, [Text.Encoding]::UTF8).TrimEnd() -split "`r?`n")
    $doc = $word.Documents.Add()
    $doc.Content.Text = ($soil -join "`r")
    $doc.Content.Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate'))
    for ($i = 0; $i -lt 900; $i++) {
        Start-Sleep -Milliseconds 100
        $working = Prop $api 'IsWorking'
        if ($i % 20 -eq 0 -or -not $working) { Step ("  working={0} status={1}" -f $working, (Prop $api 'LastStatus')) }
        if (-not $working) { break }
    }
    Step ("pictures in document: {0}" -f $doc.InlineShapes.Count)
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
