param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }
function Call($obj, [string]$name, [object[]]$arguments = @()) { [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments) }
function Show([string]$s) { $s -replace "`r", '\r' -replace "`a", '\a' -replace "`v", '\v' -replace "`n", '\n' }

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $api = $word.COMAddIns.Item('CalcpadWord.Connect').Object
    $doc = Call $word.Documents 'Open' @([string](Join-Path $OutDir 'soil-copy.docx'))
    $n = $doc.ContentControls.Count
    Step "controls: $n"
    for ($i = 1; $i -le $n; $i++) {
        $cc = $doc.ContentControls.Item($i)
        $tag = $cc.Tag
        $kind = if ($tag -like 'calcpad-block:*') { 'block' } else { 'inline' }
        $src = Show ($tag.Split(':', 4)[3])
        if ($src.Length -gt 50) { $src = $src.Substring(0, 50) + '...' }
        $hidden = $cc.Range.Font.Hidden
        $r = $cc.Range; $r.Collapse(0); $r.Select()
        try {
            [void](Call $api 'EditSource')
            $ok = 'ok'
            [void](Call $api 'Evaluate')   # back to the result
        }
        catch { $ok = 'FAILED: ' + $_.Exception.InnerException.Message }
        $para = $cc.Range.Paragraphs.Item(1).Range
        Step ("[{0,2}] {1,-6} hidden={2,-3} paraText=[{3}] src=[{4}] -> {5}" -f $i, $kind, $hidden, (Show $para.Text).Substring(0, [Math]::Min(40, (Show $para.Text).Length)), $src, $ok)
    }
    [void](Call $doc 'Close' @(0))
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { [void](Call $word 'Quit' @(0)) } catch { } }
