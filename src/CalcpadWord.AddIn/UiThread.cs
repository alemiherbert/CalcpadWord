using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace CalcpadWord.AddIn;

/// <summary>
/// Runs work on Word's UI thread, from any thread. A hidden message-only window is created on the
/// UI thread at start-up; posting to it queues work that Word's own message loop runs as soon as
/// Word is idle.
/// <para>
/// Used to defer work until Word has finished the current event or keystroke, to bring results
/// from the background thread back to Word, and to move calls from other processes (which COM
/// delivers on worker threads, because .NET objects are free-threaded) onto the UI thread.
/// Word's object model must only be used from the UI thread.
/// </para>
/// </summary>
internal static class UiThread
{
    private const uint WmRunQueued = 0x8001; // WM_APP + 1
    private const uint WmTimer = 0x0113;
    private static readonly IntPtr RunTimerId = 1;
    private static readonly IntPtr SafetyTimerId = 2;
    private const uint SafetyMilliseconds = 250;
    private static long _postedAt; // Environment.TickCount64 of the last post
    private static readonly IntPtr HwndMessage = new(-3);

    private static readonly WndProc Procedure = WindowProcedure; // kept alive: native code holds a pointer to it
    private static readonly ConcurrentQueue<Action> Queue = new();
    private static readonly HashSet<string> PendingKeys = new(StringComparer.Ordinal);
    private static readonly Lock KeysLock = new();
    private static volatile IntPtr _window;
    private static int _uiThreadId;
    private static bool _running;
    private static string? _className;

    public static bool IsUiThread => Environment.CurrentManagedThreadId == _uiThreadId;

    /// <summary>Call once, on Word's UI thread (from OnConnection).</summary>
    public static void Initialize()
    {
        _uiThreadId = Environment.CurrentManagedThreadId;
        _className = "CalcpadWord.UiThread." + Environment.ProcessId;
        var windowClass = new WndClassEx
        {
            Size = (uint)Marshal.SizeOf<WndClassEx>(),
            WindowProcedure = Marshal.GetFunctionPointerForDelegate(Procedure),
            Instance = GetModuleHandleW(null),
            ClassName = _className,
        };
        if (RegisterClassExW(ref windowClass) == 0)
            Log.Error($"RegisterClassEx failed ({Marshal.GetLastWin32Error()})");
        _window = CreateWindowExW(0, _className, string.Empty, 0, 0, 0, 0, 0, HwndMessage, IntPtr.Zero, windowClass.Instance, IntPtr.Zero);
        if (_window == IntPtr.Zero)
            Log.Error($"Creating the message window failed ({Marshal.GetLastWin32Error()}): deferred work won't run");
        else
            SetTimer(_window, SafetyTimerId, SafetyMilliseconds, IntPtr.Zero); // in case a wake-up message is lost
    }

    public static void Shutdown()
    {
        var window = _window;
        _window = IntPtr.Zero; // first: callers waiting in Invoke give up
        if (window != IntPtr.Zero)
            DestroyWindow(window); // also stops its timers
        if (_className is not null)
            UnregisterClassW(_className, GetModuleHandleW(null));
    }

    /// <summary>
    /// Runs <paramref name="action"/> on the UI thread after whatever it is doing now. With a
    /// <paramref name="key"/>, a second post with the same key is dropped while the first waits.
    /// Any thread.
    /// </summary>
    public static void Post(Action action, string? key = null)
    {
        if (key is not null)
        {
            lock (KeysLock)
                if (!PendingKeys.Add(key))
                    return;
            var inner = action;
            action = () =>
            {
                lock (KeysLock)
                    PendingKeys.Remove(key);
                inner();
            };
        }
        Queue.Enqueue(action);
        Interlocked.Exchange(ref _postedAt, Environment.TickCount64);
        if (!PostMessageW(_window, WmRunQueued, IntPtr.Zero, IntPtr.Zero))
            Log.Error($"PostMessage failed ({Marshal.GetLastWin32Error()}): deferred work is waiting");
    }

    /// <summary>Runs <paramref name="action"/> on the UI thread once <paramref name="task"/> has finished. Any thread.</summary>
    public static void WhenDone(Task task, Action action) =>
        task.ContinueWith(_ => Post(action), CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);

    /// <summary>Runs <paramref name="function"/> on the UI thread and waits for it; exceptions come back to the caller.</summary>
    public static T Invoke<T>(Func<T> function)
    {
        if (IsUiThread || _window == IntPtr.Zero)
            return function();

        var done = new ManualResetEventSlim(); // not disposed: a late run may still set it
        T result = default!;
        ExceptionDispatchInfo? error = null;
        Post(() =>
        {
            try { result = function(); }
            catch (Exception ex) { error = ExceptionDispatchInfo.Capture(ex); }
            finally { done.Set(); }
        });
        while (!done.Wait(1000))
        {
            if (_window == IntPtr.Zero)
                throw new InvalidOperationException("Word is closing.");
        }
        error?.Throw();
        return result;
    }

    private static IntPtr WindowProcedure(IntPtr window, uint message, IntPtr wParam, IntPtr lParam)
    {
        if (message == WmRunQueued)
        {
            // Not straight away: a posted message can arrive while Word is still in the middle of
            // something (a selection change, an event), when it refuses edits such as InsertXML.
            // A timer message only comes once Word's queue is empty: Word has finished.
            SetTimer(window, RunTimerId, 0, IntPtr.Zero);
            return IntPtr.Zero;
        }
        if (message != WmTimer || wParam != RunTimerId && wParam != SafetyTimerId)
            return DefWindowProcW(window, message, wParam, lParam);

        // Not again while work from this queue is running (a modal loop could dispatch the
        // message): the timer stays, so it comes back.
        if (_running)
            return IntPtr.Zero;

        if (wParam == SafetyTimerId)
        {
            // Timer messages only come when Word is idle: queued work should have run by now.
            if (Queue.IsEmpty || Environment.TickCount64 - Interlocked.Read(ref _postedAt) < SafetyMilliseconds)
                return IntPtr.Zero;
            Log.Info("Queued work was still waiting: its wake-up message was lost. Running it now.");
        }
        else
        {
            KillTimer(window, RunTimerId);
        }
        _running = true;
        try
        {
            while (Queue.TryDequeue(out var action))
            {
                Log.Trace($"UI work start: {action.Method.DeclaringType?.Name}.{action.Method.Name}");
                try { action(); }
                catch (Exception ex) { Log.Error("Deferred work failed", ex); } // never let an exception reach native code
                Log.Trace("UI work end");
            }
        }
        finally
        {
            _running = false;
        }
        return IntPtr.Zero;
    }

    private delegate IntPtr WndProc(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WndClassEx
    {
        public uint Size;
        public uint Style;
        public IntPtr WindowProcedure;
        public int ClassExtra;
        public int WindowExtra;
        public IntPtr Instance;
        public IntPtr Icon;
        public IntPtr Cursor;
        public IntPtr Background;
        public string? MenuName;
        public string ClassName;
        public IntPtr SmallIcon;
    }

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern ushort RegisterClassExW(ref WndClassEx windowClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterClassW(string className, IntPtr instance);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateWindowExW(uint exStyle, string className, string windowName, uint style,
        int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyWindow(IntPtr window);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessageW(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr DefWindowProcW(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SetTimer(IntPtr window, IntPtr id, uint milliseconds, IntPtr timerProc);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool KillTimer(IntPtr window, IntPtr id);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandleW(string? moduleName);
}
