# Samples where a thread of a 64-bit process is executing: its CPU time growth, the module of its
# instruction pointer N times, and a heuristic scan of its stack for return addresses in modules.
param([Parameter(Mandatory)][int]$ProcessId, [int]$ThreadId = 0, [int]$Samples = 25)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Sampler
{
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenThread(int access, bool inherit, int id);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(int access, bool inherit, int id);
    [DllImport("kernel32.dll")] static extern int SuspendThread(IntPtr thread);
    [DllImport("kernel32.dll")] static extern int ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool GetThreadContext(IntPtr thread, IntPtr context);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] buffer, IntPtr size, out IntPtr read);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    // Returns { rip, rsp } and fills stack with up to 64 KB from rsp.
    public static long[] Sample(int processId, int threadId, byte[] stack)
    {
        IntPtr thread = OpenThread(0x0002 | 0x0008 | 0x0040, false, threadId); // SUSPEND_RESUME | GET_CONTEXT | QUERY_INFORMATION
        IntPtr process = OpenProcess(0x0010 | 0x0400, false, processId);       // VM_READ | QUERY_INFORMATION
        if (thread == IntPtr.Zero || process == IntPtr.Zero) throw new Exception("open failed: " + Marshal.GetLastWin32Error());
        IntPtr raw = Marshal.AllocHGlobal(1232 + 16);
        IntPtr context = new IntPtr((raw.ToInt64() + 15) & ~15L);
        try
        {
            for (int i = 0; i < 1232; i++) Marshal.WriteByte(context, i, 0);
            Marshal.WriteInt32(context, 0x30, 0x100001 | 0x100002); // CONTEXT_CONTROL | CONTEXT_INTEGER
            SuspendThread(thread);
            try
            {
                if (!GetThreadContext(thread, context)) throw new Exception("GetThreadContext failed: " + Marshal.GetLastWin32Error());
                long rip = Marshal.ReadInt64(context, 0xF8), rsp = Marshal.ReadInt64(context, 0x98);
                IntPtr read;
                if (stack != null) ReadProcessMemory(process, new IntPtr(rsp), stack, new IntPtr(stack.Length), out read);
                return new long[] { rip, rsp };
            }
            finally { ResumeThread(thread); }
        }
        finally { Marshal.FreeHGlobal(raw); CloseHandle(thread); CloseHandle(process); }
    }
}
'@
$p = Get-Process -Id $ProcessId -ErrorAction Stop
if ($ThreadId -eq 0) { $ThreadId = ($p.Threads | Sort-Object StartTime | Select-Object -First 1).Id }
$modules = @($p.Modules | ForEach-Object { [pscustomobject]@{ Name = $_.ModuleName; Base = $_.BaseAddress.ToInt64(); End = $_.BaseAddress.ToInt64() + $_.ModuleMemorySize } })
function Where-Address([long]$a) { $m = $modules | Where-Object { $a -ge $_.Base -and $a -lt $_.End } | Select-Object -First 1; if ($m) { '{0}+0x{1:X}' -f $m.Name, ($a - $m.Base) } else { $null } }

$t = $p.Threads | Where-Object Id -eq $ThreadId
$cpu1 = $t.TotalProcessorTime; Start-Sleep -Seconds 3; $p.Refresh(); $t = $p.Threads | Where-Object Id -eq $ThreadId
"thread $ThreadId of $ProcessId : state $($t.ThreadState)/$($t.WaitReason); CPU used in 3 s: {0:F2} s" -f ($t.TotalProcessorTime - $cpu1).TotalSeconds

$ips = @{}
for ($i = 0; $i -lt $Samples; $i++) {
    $r = [Sampler]::Sample($ProcessId, $ThreadId, $null)
    $where = Where-Address $r[0]; if (-not $where) { $where = '0x{0:X}' -f $r[0] }
    $ips[$where] = 1 + [int]$ips[$where]
    Start-Sleep -Milliseconds 40
}
"instruction pointer, $Samples samples:"
$ips.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 12 | ForEach-Object { "  {0,3}x {1}" -f $_.Value, $_.Key }

$stack = New-Object byte[] 65536
$r = [Sampler]::Sample($ProcessId, $ThreadId, $stack)
"stack scan (return-address candidates from rsp up, first 40):"
$seen = 0
for ($o = 0; $o -le $stack.Length - 8 -and $seen -lt 40; $o += 8) {
    $v = [BitConverter]::ToInt64($stack, $o)
    $where = Where-Address $v
    if ($where -and $where -notmatch '\+0x0$') { "  rsp+0x{0:X}: {1}" -f $o, $where; $seen++ }
}
