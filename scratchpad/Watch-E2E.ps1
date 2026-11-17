# Runs the end-to-end test several times. If, while the test's Word is alive and the test hasn't
# finished, its log stops growing for 45 s, dumps that Word's managed stacks to stacks-<n>.txt.
param([int]$Runs = 4)
$root = 'c:\Users\dell alienware\Desktop\Alemi Herbert\Plugin'
$out = 'C:\Users\DELLAL~1\AppData\Local\Temp\claude\2df21e2bc58c4afc\1a257411-5f73-49ec-b568-ebeb1593d5ab\scratchpad\watch'
New-Item -ItemType Directory -Force $out | Out-Null
$stack = "$env:USERPROFILE\.dotnet\tools\dotnet-stack.exe"
$log = Join-Path $root 'out\e2e\e2e.log'
$pidFile = Join-Path $root 'out\e2e\word.pid'
for ($n = 1; $n -le $Runs; $n++) {
    Remove-Item $pidFile -ErrorAction SilentlyContinue
    $job = Start-Job -ScriptBlock { param($r) & (Join-Path $r 'tools\Test-InWord.ps1') -TimeoutSeconds 400 2>&1 } -ArgumentList $root
    $lastSize = -1; $quietSince = Get-Date; $dumped = $false
    while ($job.State -eq 'Running') {
        Start-Sleep -Seconds 2
        $size = if (Test-Path $log) { (Get-Item $log).Length } else { 0 }
        if ($size -ne $lastSize) { $lastSize = $size; $quietSince = Get-Date; continue }
        $wordAlive = (Test-Path $pidFile) -and (Get-Process -Id ([int](Get-Content $pidFile)) -ErrorAction SilentlyContinue)
        $finished = (Test-Path $log) -and (Select-String -Path $log -Pattern 'ALL PASSED|FAILED$' -Quiet -ErrorAction SilentlyContinue)
        if (-not $dumped -and $wordAlive -and -not $finished -and ((Get-Date) - $quietSince).TotalSeconds -gt 45) {
            $wpid = [int](Get-Content $pidFile)
            "run ${n}: quiet for 45 s; dumping Word $wpid"
            & $stack report -p $wpid 2>&1 | Out-File -Encoding utf8 (Join-Path $out "stacks-$n.txt")
            & (Join-Path (Split-Path $out -Parent) 'Get-WaitChain.ps1') -ProcessId $wpid 2>&1 | Out-File -Encoding utf8 (Join-Path $out "waitchain-$n.txt")
            Copy-Item $log (Join-Path $out "e2e-$n.log")
            $dumped = $true
        }
    }
    $result = @(Receive-Job $job -ErrorAction SilentlyContinue); Remove-Job $job -Force
    "run ${n}: " + (($result | Where-Object { $_ -match 'ALL PASSED|FAILED$|TIMEOUT' }) -join ' ')
    if ($dumped) { break }
}
"lost wake-ups logged: " + @(Select-String -Path "$env:LOCALAPPDATA\CalcpadWord\addin.log" -Pattern 'wake-up message was lost').Count
