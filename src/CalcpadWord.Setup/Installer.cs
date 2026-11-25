using System.Diagnostics;
using System.IO.Compression;
using Microsoft.Win32;

namespace CalcpadWord.Setup;

/// <summary>
/// Per-user install: files under %LOCALAPPDATA%\Programs\CalcpadWord (the add-in, Calcpad's engine,
/// a private .NET runtime and the fonts), fonts in the user's font folder, registration under HKCU
/// only (COM class, Word add-in entry, and an entry in Settings > Apps for uninstalling).
/// Each install goes into its own app-* folder, so it works while Word has an older version
/// loaded; older folders are removed once no Word is using them.
/// </summary>
internal static class Installer
{
    private static readonly string Root =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "CalcpadWord");
    private static readonly string InstalledSetup = Path.Combine(Root, "CalcpadWord-Setup.exe");

    private const string ClassesKey = @"Software\Classes";
    private const string WordAddinsKey = @"Software\Microsoft\Office\Word\Addins";
    private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\CalcpadWord";

    public static int Install(Ui ui)
    {
        ui.Log($"Install {Program.Version} for {Environment.UserName}");
        if (!Environment.Is64BitOperatingSystem)
        {
            ui.Error($"{AddInIdentity.Name} needs 64-bit Windows.");
            return 2;
        }

        var word = WordInstallation.Detect();
        ui.Log($"Word: {word.Description}");
        if (word.Found && !word.Is64Bit)
        {
            ui.Error($"{AddInIdentity.Name} works with 64-bit Office only; this computer has {word.Description}.");
            return 2;
        }
        if (!word.Found && !ui.Confirm("Microsoft Word was not found on this computer. Install anyway?"))
            return 2;

        if (!ui.Confirm(
                $"Install {AddInIdentity.Name} {Program.Version}?\n\n" +
                $"It is installed for {Environment.UserName} only, without administrator rights. It adds a " +
                "Calcpad tab to Word's ribbon and installs its fonts (math fonts and Libre Franklin). " +
                "Everything it needs is included: Calcpad and .NET don't have to be installed.\n\n" +
                "You can remove it later in Settings > Apps."))
            return 2;

        using var payload = typeof(Installer).Assembly.GetManifestResourceStream("payload.zip")
            ?? throw new InvalidOperationException("This setup was built without the add-in. Build it with tools\\Build-Installer.ps1.");

        var app = Path.Combine(Root, $"app-{Program.Version}-{DateTime.Now:yyyyMMddHHmmss}");
        Directory.CreateDirectory(app);
        ZipFile.ExtractToDirectory(payload, app);
        var loader = Path.Combine(app, AddInIdentity.LoaderFile);
        foreach (var required in new[] { loader, Path.Combine(app, AddInIdentity.ComHostFile), Path.Combine(app, "runtime", "host", "fxr") })
        {
            if (!File.Exists(required) && !Directory.Exists(required))
                throw new FileNotFoundException("This setup is incomplete: rebuild it with tools\\Build-Installer.ps1.", required);
        }
        ui.Log($"Files in {app}");

        // Keep a copy of this setup for uninstalling (unless it is that copy).
        var self = Environment.ProcessPath!;
        if (!string.Equals(Path.GetFullPath(self), InstalledSetup, StringComparison.OrdinalIgnoreCase))
            File.Copy(self, InstalledSetup, overwrite: true);

        var fontProblems = InstallFonts(Path.Combine(app, "fonts"), ui);
        Register(loader, DirectorySize(Root));
        var kept = RemoveOldVersions(keep: app);

        var done = WordIsRunning()
            ? $"{AddInIdentity.Name} {Program.Version} is installed.\n\nIf Word is open, restart it to load this version; the Calcpad tab is on the ribbon."
            : $"{AddInIdentity.Name} {Program.Version} is installed.\n\nStart Word: the Calcpad tab is on the ribbon.";
        if (fontProblems > 0)
            ui.Warning($"{done}\n\n{fontProblems} font(s) couldn't be installed. Details: {Ui.LogPath}");
        else
            ui.Info(done);
        if (kept > 0)
            ui.Log($"{kept} older version(s) are still in use by Word; they are removed on the next install or uninstall.");
        return 0;
    }

    /// <summary>
    /// Installs the bundled fonts for this user (fonts\fonts.txt lists each file with its registry
    /// name); fonts installed for all users are left alone. Returns how many failed.
    /// </summary>
    private static int InstallFonts(string folder, Ui ui)
    {
        var list = Path.Combine(folder, "fonts.txt");
        if (!File.Exists(list))
            return 0;

        int installed = 0, present = 0, failed = 0;
        foreach (var line in File.ReadAllLines(list))
        {
            var parts = line.Split('\t');
            if (parts.Length != 2)
                continue;
            var (file, name) = (parts[0], parts[1]);
            try
            {
                if (UserFonts.IsInstalledForAllUsers(name))
                {
                    ++present;
                    continue;
                }
                var (path, loaded) = UserFonts.Install(file, File.ReadAllBytes(Path.Combine(folder, file)), name);
                ++installed;
                if (!loaded)
                    ui.Log($"Font {name}: installed at {path}; Windows loads it at the next sign-in.");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                ++failed;
                ui.Log($"Font {name}: {ex.Message}");
            }
        }
        UserFonts.Announce();
        ui.Log($"Fonts: {installed} installed for this user, {present} already installed for all users, {failed} failed.");
        return failed;
    }

    public static int Uninstall(Ui ui)
    {
        ui.Log($"Uninstall {Program.Version}");
        if (!ui.Confirm($"Remove {AddInIdentity.Name}?\n\nDocuments keep their last calculated results. The fonts it installed stay installed."))
            return 2;

        Unregister();

        var locked = RemoveOldVersions(keep: null);
        var runningFromRoot = Path.GetFullPath(Environment.ProcessPath!)
            .StartsWith(Root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        if (runningFromRoot)
            DeleteAfterExit(Root);   // this exe can't delete itself while it runs
        else if (locked == 0)
            TryDelete(Root);

        ui.Info(locked > 0
            ? $"{AddInIdentity.Name} is removed. Some of its files are in use by Word: close Word to let them go."
            : WordIsRunning()
                ? $"{AddInIdentity.Name} is removed. If Word is open, restart it to remove the Calcpad tab."
                : $"{AddInIdentity.Name} is removed.");
        return 0;
    }

    // ---- Registry -----------------------------------------------------------------------

    /// <param name="server">The loader DLL: Word loads it, it starts the bundled .NET and hands over to the COM host.</param>
    private static void Register(string server, long sizeBytes)
    {
        var clsid = "{" + AddInIdentity.ClassId + "}";
        var hkcu = Registry.CurrentUser; // shared system key: not disposed

        using (var key = hkcu.CreateSubKey($@"{ClassesKey}\CLSID\{clsid}"))
            key.SetValue(null, AddInIdentity.Name);
        using (var key = hkcu.CreateSubKey($@"{ClassesKey}\CLSID\{clsid}\InprocServer32"))
        {
            key.SetValue(null, server);
            key.SetValue("ThreadingModel", "Both");
        }
        using (var key = hkcu.CreateSubKey($@"{ClassesKey}\CLSID\{clsid}\ProgID"))
            key.SetValue(null, AddInIdentity.ProgId);
        using (var key = hkcu.CreateSubKey($@"{ClassesKey}\{AddInIdentity.ProgId}"))
            key.SetValue(null, AddInIdentity.Name);
        using (var key = hkcu.CreateSubKey($@"{ClassesKey}\{AddInIdentity.ProgId}\CLSID"))
            key.SetValue(null, clsid);

        using (var key = hkcu.CreateSubKey($@"{WordAddinsKey}\{AddInIdentity.ProgId}"))
        {
            key.SetValue("FriendlyName", AddInIdentity.Name);
            key.SetValue("Description", AddInIdentity.Description);
            key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); // load at startup
        }

        using (var key = hkcu.CreateSubKey(UninstallKey))
        {
            key.SetValue("DisplayName", AddInIdentity.Name);
            key.SetValue("DisplayVersion", Program.Version);
            key.SetValue("DisplayIcon", InstalledSetup);
            key.SetValue("InstallLocation", Root);
            key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            key.SetValue("UninstallString", $"\"{InstalledSetup}\" /uninstall");
            key.SetValue("QuietUninstallString", $"\"{InstalledSetup}\" /uninstall /quiet");
            key.SetValue("EstimatedSize", (int)(sizeBytes / 1024), RegistryValueKind.DWord);
            key.SetValue("NoModify", 1, RegistryValueKind.DWord);
            key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        }
    }

    private static void Unregister()
    {
        var clsid = "{" + AddInIdentity.ClassId + "}";
        var hkcu = Registry.CurrentUser; // shared system key: not disposed
        hkcu.DeleteSubKeyTree($@"{WordAddinsKey}\{AddInIdentity.ProgId}", throwOnMissingSubKey: false);
        hkcu.DeleteSubKeyTree($@"{ClassesKey}\CLSID\{clsid}", throwOnMissingSubKey: false);
        hkcu.DeleteSubKeyTree($@"{ClassesKey}\{AddInIdentity.ProgId}", throwOnMissingSubKey: false);
        hkcu.DeleteSubKeyTree(UninstallKey, throwOnMissingSubKey: false);
    }

    // ---- Files --------------------------------------------------------------------------

    /// <summary>Deletes app-* folders other than <paramref name="keep"/>; returns how many are still in use.</summary>
    private static int RemoveOldVersions(string? keep)
    {
        if (!Directory.Exists(Root))
            return 0;

        foreach (var dir in Directory.GetDirectories(Root, $"{RemovingPrefix}*"))
            TryDelete(dir); // left by an earlier install

        var inUse = 0;
        foreach (var dir in Directory.GetDirectories(Root, "app-*"))
        {
            if (keep is not null && string.Equals(dir, keep, StringComparison.OrdinalIgnoreCase))
                continue;
            if (!TryRemoveVersion(dir))
                ++inUse;
        }
        return inUse;
    }

    private const string RemovingPrefix = "removing-";

    /// <summary>
    /// Deletes a version's folder only if no Word is using it. Deleting it file by file would take
    /// the files a running Word hasn't loaded yet (SkiaSharp, loaded at the first plot) and break
    /// its add-in, so the folder is renamed first: Windows refuses that while a file in it is open.
    /// </summary>
    private static bool TryRemoveVersion(string dir)
    {
        var removing = Path.Combine(Root, RemovingPrefix + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.Move(dir, removing);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false; // a running Word has the add-in's files open
        }
        TryDelete(removing); // anything left goes on the next install or uninstall
        return true;
    }

    private static bool TryDelete(string dir)
    {
        try
        {
            Directory.Delete(dir, recursive: true);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false; // a running Word has the add-in's files open
        }
    }

    private static void DeleteAfterExit(string dir)
    {
        var temp = Path.GetTempPath();
        Environment.CurrentDirectory = temp;
        Process.Start(new ProcessStartInfo
        {
            FileName = Path.Combine(Environment.SystemDirectory, "cmd.exe"),
            Arguments = $"/c ping 127.0.0.1 -n 3 > nul & rmdir /s /q \"{dir}\"",
            WorkingDirectory = temp,
            CreateNoWindow = true,
            UseShellExecute = false,
        });
    }

    private static long DirectorySize(string dir) =>
        Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length);

    private static bool WordIsRunning() => Process.GetProcessesByName("WINWORD").Length > 0;
}

/// <summary>Which Word is installed: the add-in is 64-bit, so 32-bit Office can't load it.</summary>
internal sealed record WordInstallation(bool Found, bool Is64Bit, string Description)
{
    public static WordInstallation Detect()
    {
        // Microsoft 365 and Office 2019/2021/2024 (Click-to-Run).
        using (var c2r = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
        {
            if (c2r?.GetValue("Platform") is string platform)
                return new(true, platform.Equals("x64", StringComparison.OrdinalIgnoreCase), $"Office ({platform}, Click-to-Run)");
        }

        // Office installed with MSI (2013, 2016).
        using var machine64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var machine32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
        foreach (var version in new[] { "16.0", "15.0" })
        {
            var path = $@"SOFTWARE\Microsoft\Office\{version}\Word\InstallRoot";
            using (var key64 = machine64.OpenSubKey(path))
                if (key64 is not null)
                    return new(true, true, $"Office {version} (64-bit)");
            using (var key32 = machine32.OpenSubKey(path))
                if (key32 is not null)
                    return new(true, false, $"Office {version} (32-bit)");
        }
        return new(false, false, "Word not found");
    }
}
