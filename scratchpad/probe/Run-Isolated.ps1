# Runs a Word automation script in a separate powershell.exe. The script must write its Word PID
# to <OutDir>\pid.txt and its progress to <OutDir>\steps.log. On timeout, kills ONLY that Word
# PID and the child shell, never the user's own Word.
param([Parameter(Mandatory)][string]$Script, [Parameter(Mandatory)][string]$OutDir, [int]$TimeoutSeconds = 120)
New-Item -ItemType Directory -Force $OutDir | Out-Null
Remove-Item (Join-Path $OutDir 'pid.txt'), (Join-Path $OutDir 'steps.log') -ErrorAction SilentlyContinue
$proc = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$Script`"", '-OutDir', "`"$OutDir`"")
if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
    "TIMEOUT after $TimeoutSeconds s"
    $pidFile = Join-Path $OutDir 'pid.txt'
    if (Test-Path $pidFile) {
        $wpid = [int](Get-Content $pidFile)
        Stop-Process -Id $wpid -Force -ErrorAction SilentlyContinue
        "Killed automation Word PID $wpid"
    }
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
}
else { "Exit code: $($proc.ExitCode)" }
$log = Join-Path $OutDir 'steps.log'
if (Test-Path $log) { Get-Content $log -Encoding UTF8 }
