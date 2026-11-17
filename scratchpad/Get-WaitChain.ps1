# Prints the wait chain (Windows Wait Chain Traversal) of every thread of a process that is
# blocked: what each one waits for, and on which other thread/process.
param([Parameter(Mandatory)][int]$ProcessId)
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class Wct
{
    [DllImport("advapi32.dll", SetLastError = true)] static extern IntPtr OpenThreadWaitChainSession(int flags, IntPtr callback);
    [DllImport("advapi32.dll")] static extern void CloseThreadWaitChainSession(IntPtr session);
    [DllImport("advapi32.dll", SetLastError = true)]
    static extern bool GetThreadWaitChain(IntPtr session, IntPtr context, int flags, int threadId, ref int nodeCount, IntPtr nodes, out bool isCycle);
    const int NodeSize = 288, MaxNodes = 16;
    static readonly string[] Types = { "?", "CriticalSection", "SendMessage", "Mutex", "Alpc", "Com", "ThreadWait", "ProcessWait", "Thread", "ComActivation", "Unknown", "SocketIo", "SmbIo" };
    static readonly string[] States = { "?", "NoAccess", "Running", "Blocked", "PidOnly", "PidOnlyRpcss", "Owned", "NotOwned", "Abandoned", "Unknown", "Error" };
    public static string Chain(int threadId)
    {
        var session = OpenThreadWaitChainSession(0, IntPtr.Zero);
        if (session == IntPtr.Zero) return "OpenThreadWaitChainSession failed: " + Marshal.GetLastWin32Error();
        var buffer = Marshal.AllocHGlobal(NodeSize * MaxNodes);
        try
        {
            int count = MaxNodes; bool cycle;
            if (!GetThreadWaitChain(session, IntPtr.Zero, 0x1 | 0x4, threadId, ref count, buffer, out cycle))
                return "GetThreadWaitChain failed: " + Marshal.GetLastWin32Error();
            var sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                IntPtr node = IntPtr.Add(buffer, i * NodeSize);
                int type = Marshal.ReadInt32(node), state = Marshal.ReadInt32(node, 4);
                string t = type < Types.Length ? Types[type] : type.ToString(), s = state < States.Length ? States[state] : state.ToString();
                if (type == 8) // thread
                    sb.AppendFormat(" -> Thread(pid {0}, tid {1}, {2}, wait {3} ms)", Marshal.ReadInt32(node, 8), Marshal.ReadInt32(node, 12), s, Marshal.ReadInt32(node, 16));
                else
                {
                    string name = Marshal.PtrToStringUni(IntPtr.Add(node, 8));
                    sb.AppendFormat(" -> {0}({1}{2})", t, s, string.IsNullOrEmpty(name) ? "" : ", " + name);
                }
            }
            if (cycle) sb.Append("  [DEADLOCK CYCLE]");
            return sb.ToString();
        }
        finally { Marshal.FreeHGlobal(buffer); CloseThreadWaitChainSession(session); }
    }
}
'@
$p = Get-Process -Id $ProcessId -ErrorAction Stop
$main = $p.Threads | Sort-Object StartTime | Select-Object -First 1
"Process $ProcessId ($($p.ProcessName)), main thread $($main.Id)"
foreach ($t in $p.Threads) {
    $chain = [Wct]::Chain($t.Id)
    $isMain = $t.Id -eq $main.Id
    if ($isMain -or $chain -match 'SendMessage|Alpc|Com|CriticalSection|Mutex|ProcessWait|DEADLOCK|pid (?!' + $ProcessId + '\b)') {
        "$(if ($isMain) { 'MAIN ' } else { '     ' })thread $($t.Id) [$($t.ThreadState)/$($t.WaitReason)]:$chain"
    }
}
