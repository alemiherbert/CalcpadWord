using System.Runtime.InteropServices;
using System.Text;

namespace CalcpadWord.AddIn;

/// <summary>A key plus modifiers, e.g. Alt+Shift+Enter.</summary>
internal readonly record struct KeyChord(int VirtualKey, bool Ctrl, bool Alt, bool Shift, string Name);

/// <param name="Claims">
/// Null for a shortcut that Word doesn't use: the key is passed on as well. Otherwise the key is
/// taken from Word only while this returns true; it must be cheap and must not call Word.
/// </param>
internal sealed record KeyBinding(KeyChord Chord, Action Action, Func<bool>? Claims = null);

/// <summary>
/// Keyboard shortcuts for a COM add-in. Word's own key bindings can only call VBA macros, so this
/// uses a keyboard hook on Word's UI thread (not a system-wide hook). It only reacts while the
/// document has the focus. Actions run right after the keystroke, through <see cref="UiThread"/>.
/// </summary>
internal sealed class KeyboardShortcuts : IDisposable
{
    public const int VkReturn = 0x0D, VkEscape = 0x1B, VkF8 = 0x77, VkF9 = 0x78;

    public static readonly KeyChord Evaluate = new(VkReturn, Ctrl: false, Alt: true, Shift: true, "Alt+Shift+Enter");
    public static readonly KeyChord Recalculate = new(VkF9, Ctrl: true, Alt: true, Shift: false, "Ctrl+Alt+F9");
    public static readonly KeyChord NextError = new(VkF8, Ctrl: true, Alt: true, Shift: false, "Ctrl+Alt+F8");
    public static readonly KeyChord Enter = new(VkReturn, Ctrl: false, Alt: false, Shift: false, "Enter");
    public static readonly KeyChord Escape = new(VkEscape, Ctrl: false, Alt: false, Shift: false, "Esc");

    private const int WhKeyboard = 2;
    private const int HcAction = 0;
    private const int VkShift = 0x10, VkControl = 0x11, VkMenu = 0x12;
    private const string DocumentWindowClass = "_WwG"; // Word's document pane

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    private readonly HookProc _proc;   // kept alive while the hook is installed
    private readonly IReadOnlyList<KeyBinding> _bindings;
    private IntPtr _hook;

    public KeyboardShortcuts(IReadOnlyList<KeyBinding> bindings)
    {
        _bindings = bindings;
        _proc = OnKey;
        _hook = SetWindowsHookExW(WhKeyboard, _proc, IntPtr.Zero, GetCurrentThreadId());
        if (_hook == IntPtr.Zero)
            Log.Error($"Keyboard shortcuts unavailable: SetWindowsHookEx failed ({Marshal.GetLastWin32Error()})");
    }

    public void Dispose()
    {
        if (_hook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }

    private IntPtr OnKey(int code, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            var flags = lParam.ToInt64();
            var isKeyDown = (flags & 0x80000000L) == 0;
            var isRepeat = (flags & 0x40000000L) != 0;
            if (code == HcAction && isKeyDown)
            {
                var vk = wParam.ToInt32();
                foreach (var binding in _bindings)
                {
                    var chord = binding.Chord;
                    if (chord.VirtualKey != vk ||
                        chord.Ctrl != IsDown(VkControl) ||
                        chord.Alt != IsDown(VkMenu) ||
                        chord.Shift != IsDown(VkShift) ||
                        !DocumentHasFocus())
                        continue;

                    if (binding.Claims is null)
                    {
                        if (!isRepeat)
                            UiThread.Post(binding.Action, key: "shortcut:" + chord.Name);
                        break; // Word doesn't use this chord; pass it on
                    }
                    if (binding.Claims())
                    {
                        if (!isRepeat)
                            UiThread.Post(binding.Action, key: "shortcut:" + chord.Name);
                        return 1; // taken: Word never sees this key
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error("Keyboard hook failed", ex);
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private static bool IsDown(int vk) => (GetKeyState(vk) & 0x8000) != 0;

    private static bool DocumentHasFocus()
    {
        var focus = GetFocus();
        if (focus == IntPtr.Zero)
            return false;

        var name = new StringBuilder(64);
        return GetClassNameW(focus, name, name.Capacity) > 0 && name.ToString() == DocumentWindowClass;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookExW(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);

    [DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
