# Runs a Word automation script in a background job; if it hangs, kills ONLY the Word
# process that script started (its PID is written to pid.txt), never the user's Word.
param([Parameter(Mandatory)][string]$Script, [Parameter(Mandatory)][string]$OutDir, [int]$TimeoutSeconds = 180)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$pidFile = Join-Path $OutDir 'pid.txt'
Remove-Item $pidFile -ErrorAction SilentlyContinue
$job = Start-Job -ScriptBlock { param($s, $o) & $s -OutDir $o 2>&1 | ForEach-Object { "$_" } } -ArgumentList $Script, $OutDir
if (-not (Wait-Job $job -Timeout $TimeoutSeconds)) {
    "TIMEOUT after $TimeoutSeconds s"
    if (Test-Path $pidFile) {
        $wpid = [int](Get-Content $pidFile)
        Stop-Process -Id $wpid -Force -ErrorAction SilentlyContinue
        "Killed automation Word PID $wpid"
    }
    Stop-Job $job
}
Receive-Job $job
Remove-Job $job -Force
