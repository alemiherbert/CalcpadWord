using System.Runtime.InteropServices;

namespace CalcpadWord.Setup;

/// <summary>Message boxes (none in quiet mode) and a log in %TEMP%.</summary>
internal sealed class Ui(bool quiet)
{
    public static readonly string LogPath = Path.Combine(Path.GetTempPath(), "CalcpadWord-Setup.log");
    private const string Caption = AddInIdentity.Name + " Setup";

    private const uint MbOk = 0x0, MbYesNo = 0x4, MbIconError = 0x10, MbIconQuestion = 0x20,
        MbIconWarning = 0x30, MbIconInformation = 0x40, MbSetForeground = 0x10000;
    private const int IdYes = 6;

    public bool Quiet { get; } = quiet;

    /// <summary>Asks a yes/no question; quiet mode answers yes.</summary>
    public bool Confirm(string question)
    {
        Log("? " + question.ReplaceLineEndings(" "));
        return Quiet || MessageBoxW(IntPtr.Zero, question, Caption, MbYesNo | MbIconQuestion | MbSetForeground) == IdYes;
    }

    public void Info(string text) => Show(text, MbIconInformation);

    public void Warning(string text) => Show(text, MbIconWarning);

    public void Error(string text) => Show(text, MbIconError);

    public void Log(string text)
    {
        try { File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {text}{Environment.NewLine}"); }
        catch { /* logging must not stop setup */ }
    }

    private void Show(string text, uint icon)
    {
        Log(text.ReplaceLineEndings(" "));
        if (!Quiet)
            _ = MessageBoxW(IntPtr.Zero, text, Caption, MbOk | icon | MbSetForeground);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);
}
