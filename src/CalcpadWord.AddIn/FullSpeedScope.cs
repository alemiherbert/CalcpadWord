using System.Runtime.InteropServices;

namespace CalcpadWord.AddIn;

/// <summary>
/// Lets the current thread run at full speed while the scope is open. When Word is not the
/// foreground app (or is hidden), Windows 11 may give its threads "Eco" quality of service:
/// efficiency cores at a low clock, which made calculations several times slower. A calculation
/// is work the user is waiting for, so it opts out. The thread's own setting is restored at the
/// end (it is a thread-pool thread that will run other work).
/// </summary>
internal sealed class FullSpeedScope : IDisposable
{
    private const int ThreadPowerThrottling = 3; // THREAD_INFORMATION_CLASS
    private const uint CurrentVersion = 1;
    private const uint ExecutionSpeed = 0x1;

    [ThreadStatic] private static int _depth; // scopes nest: only the outermost one changes the setting

    private readonly bool _changed;
    private bool _disposed;

    public FullSpeedScope()
    {
        // ControlMask: we decide execution speed. StateMask clear: not throttled.
        if (_depth++ == 0)
            _changed = Set(ExecutionSpeed, 0);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        if (--_depth == 0 && _changed)
            Set(0, 0); // back to letting Windows decide
    }

    private static bool Set(uint controlMask, uint stateMask)
    {
        var state = new ThrottlingState { Version = CurrentVersion, ControlMask = controlMask, StateMask = stateMask };
        try
        {
            return SetThreadInformation(GetCurrentThread(), ThreadPowerThrottling, ref state, (uint)Marshal.SizeOf<ThrottlingState>());
        }
        catch (EntryPointNotFoundException)
        {
            return false; // Windows before 8: no such setting
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ThrottlingState
    {
        public uint Version;
        public uint ControlMask;
        public uint StateMask;
    }

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentThread();

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetThreadInformation(IntPtr thread, int informationClass, ref ThrottlingState information, uint size);
}
