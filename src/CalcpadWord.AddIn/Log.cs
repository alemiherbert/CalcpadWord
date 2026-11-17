namespace CalcpadWord.AddIn;

/// <summary>Append-only log at %LOCALAPPDATA%\CalcpadWord\addin.log. Never throws.</summary>
internal static class Log
{
    public static readonly string FilePath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CalcpadWord", "addin.log");

    private static readonly Lock Gate = new();

    public static void Info(string message) => Write("INFO ", message);

    /// <summary>
    /// Detailed tracing for finding hangs: on only while %LOCALAPPDATA%\CalcpadWord\trace.on
    /// exists when Word starts. Each line has the thread it ran on.
    /// </summary>
    public static readonly bool Tracing = File.Exists(Path.Combine(Path.GetDirectoryName(FilePath)!, "trace.on"));

    public static void Trace(string message)
    {
        if (Tracing)
            Write("TRACE", $"[{Environment.CurrentManagedThreadId}] {message}");
    }

    public static void Error(string message, Exception? ex = null) =>
        Write("ERROR", ex is null ? message : $"{message}: {ex}");

    private static void Write(string level, string message)
    {
        try
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                File.AppendAllText(FilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {level} {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Logging must never break Word.
        }
    }
}

/// <summary>A problem the user can fix (no document open, cursor in the wrong place, ...).</summary>
internal sealed class UserMessageException(string message) : Exception(message);
