using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace CalcpadWord;

/// <summary>
/// Installs fonts for the current user, the way Windows does when you choose "Install" (not
/// "Install for all users"): the file goes to %LOCALAPPDATA%\Microsoft\Windows\Fonts and is
/// registered under HKCU. No administrator rights. Compiled into both the add-in and Setup.
/// </summary>
internal static class UserFonts
{
    public static readonly string Folder =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Fonts");

    public const string RegistryKey = @"Software\Microsoft\Windows NT\CurrentVersion\Fonts";

    /// <summary>The registry name Windows gives a font: "Family Style (OpenType)" or "(TrueType)".</summary>
    public static string RegistryName(string fullName, byte[] data) =>
        $"{fullName} ({(data.Length > 0 && data[0] == 'O' ? "OpenType" : "TrueType")})";

    /// <summary>True if this font is installed for all users (then a per-user copy isn't needed).</summary>
    public static bool IsInstalledForAllUsers(string registryName)
    {
        using var key = Registry.LocalMachine.OpenSubKey(RegistryKey);
        if (key?.GetValue(registryName) is not string file || file.Length == 0)
            return false;
        var path = Path.IsPathRooted(file) ? file : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), file);
        return File.Exists(path);
    }

    /// <summary>
    /// Copies the font to the user's font folder (unless the same file is there already), registers
    /// it and adds it to this session. Returns the installed file and whether Windows loaded it.
    /// Call <see cref="Announce"/> once afterwards so running programs see the new fonts.
    /// </summary>
    public static (string File, bool Loaded) Install(string fileName, byte[] data, string registryName)
    {
        Directory.CreateDirectory(Folder);
        var file = Path.Combine(Folder, Path.GetFileName(fileName));
        if (!File.Exists(file) || new FileInfo(file).Length != data.Length)
            File.WriteAllBytes(file, data);

        using (var key = Registry.CurrentUser.CreateSubKey(RegistryKey))
            key.SetValue(registryName, file);

        // Usable now, without signing out.
        return (file, AddFontResourceW(file) != 0);
    }

    /// <summary>Tells running programs (Word among them) that the fonts changed.</summary>
    public static void Announce() =>
        SendMessageTimeoutW(HwndBroadcast, WmFontChange, IntPtr.Zero, IntPtr.Zero, SmtoAbortIfHung, 2000, out _);

    private static readonly IntPtr HwndBroadcast = new(0xFFFF);
    private const uint WmFontChange = 0x001D;
    private const uint SmtoAbortIfHung = 0x0002;

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern int AddFontResourceW(string fileName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
}
