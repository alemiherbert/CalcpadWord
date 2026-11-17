param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Show([string]$s) { $s -replace "`r", '\r' -replace "`a", '\a' -replace "`v", '\v' }
function Dump($doc, [string]$label) {
    Step "$label : $($doc.Paragraphs.Count) paragraphs"
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $cc = $p.Range.Characters.Item(1).ParentContentControl
        $t = Show $p.Range.Text
        if ($t.Length -gt 70) { $t = $t.Substring(0, 70) + '...' }
        Step ("  [{0,2}] outline={1,-2} cc={2,-14} [{3}]" -f $i, $p.OutlineLevel, $(if ($cc) { $cc.Tag.Substring(0, [Math]::Min(14, $cc.Tag.Length)) } else { '-' }), $t)
    }
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object

    $doc = $word.Documents.Add()
    $doc.Content.Text = @('"Title', "'Text 'x = 2", 'x*3', 'After.') -join "`r"
    $doc.Range(0, $doc.Paragraphs.Item(3).Range.End).Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate'))
    Dump $doc 'simple conversion'
    Call $doc 'Close' @(0) | Out-Null

    $doc = $word.Documents.Add()
    $doc.Content.Text = @("'Text 'x = 2", '$Plot{t @ t = 0 : x}', 'After.') -join "`r"
    $doc.Range(0, $doc.Paragraphs.Item(2).Range.End).Select()
    Step ('Evaluate -> ' + (Call $api 'Evaluate'))
    Dump $doc 'conversion ending in a block'
    Call $doc 'Close' @(0) | Out-Null
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { Call $word 'Quit' @(0) | Out-Null } catch { } }
