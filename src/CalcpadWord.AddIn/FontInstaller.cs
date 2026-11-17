using CalcpadWord.Engine;
using Microsoft.Win32;

namespace CalcpadWord.AddIn;

/// <summary>
/// Installs a free math font for the current user (see <see cref="UserFonts"/>): no administrator
/// rights. Setup installs the bundled fonts already; this reinstalls one that was removed, from the
/// copy bundled with the add-in, or downloads it. The file is checked: it must be a font file
/// containing the expected family, with an OpenType MATH table.
/// </summary>
internal static class FontInstaller
{
    private static readonly HttpClient Http = CreateClient();

    private const string FontsKey = UserFonts.RegistryKey;

    /// <summary>The fonts bundled with the add-in (Setup installs them).</summary>
    private static readonly string BundledFolder =
        Path.Combine(Path.GetDirectoryName(typeof(FontInstaller).Assembly.Location) ?? ".", "fonts");

    /// <summary>Installs the font from the bundled copy, or downloads it (runs on a background thread). Returns the installed file.</summary>
    public static async Task<string> InstallAsync(MathFontInfo font, CancellationToken cancellationToken = default)
    {
        if (font.DownloadUrl is null)
            throw new InvalidOperationException($"{font.Family} can't be downloaded: {font.Note}");

        var fileName = Path.GetFileName(new Uri(font.DownloadUrl).LocalPath);
        var bundled = Path.Combine(BundledFolder, fileName);
        var source = File.Exists(bundled) ? bundled : font.DownloadUrl;
        var data = File.Exists(bundled)
            ? await File.ReadAllBytesAsync(bundled, cancellationToken).ConfigureAwait(false)
            : await Http.GetByteArrayAsync(font.DownloadUrl, cancellationToken).ConfigureAwait(false);
        var contents = FontFile.Read(data);
        if (!contents.Any(f => f.Family.Equals(font.Family, StringComparison.OrdinalIgnoreCase) && f.IsMath))
            throw new InvalidDataException($"The file from {source} is not the {font.Family} math font.");
        cancellationToken.ThrowIfCancellationRequested();

        var (file, loaded) = UserFonts.Install(fileName, data, UserFonts.RegistryName(font.Family, data));
        if (!loaded)
            Log.Error($"AddFontResource failed for {file}");
        UserFonts.Announce();
        Log.Info($"Installed {font.Family} for the current user from {source}: {file}");
        return file;
    }

    /// <summary>The font files registered with Windows (for all users, then this user), as full paths.</summary>
    public static IReadOnlyList<string> RegisteredFiles()
    {
        var systemFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        var files = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var hive in new[] { Registry.LocalMachine, Registry.CurrentUser })
        {
            using var key = hive.OpenSubKey(FontsKey);
            if (key is null)
                continue;
            foreach (var name in key.GetValueNames())
            {
                if (key.GetValue(name) is not string file || file.Length == 0)
                    continue;
                var path = Path.IsPathRooted(file) ? file : Path.Combine(systemFonts, file);
                if (seen.Add(path))
                    files.Add(path);
            }
        }
        return files;
    }

    /// <summary>Font families registered with Windows (for all users or this user).</summary>
    public static HashSet<string> RegisteredFamilies()
    {
        var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var hive in new[] { Registry.LocalMachine, Registry.CurrentUser })
        {
            using var key = hive.OpenSubKey(FontsKey);
            if (key is null)
                continue;
            foreach (var name in key.GetValueNames())
            {
                // "Cambria & Cambria Math (TrueType)" → Cambria, Cambria Math
                var bare = name.Contains(" (") ? name[..name.LastIndexOf(" (", StringComparison.Ordinal)] : name;
                foreach (var family in bare.Split(" & ", StringSplitOptions.TrimEntries))
                    families.Add(family);
            }
        }
        return families;
    }

    private static HttpClient CreateClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("CalcpadForWord-FontInstaller/1.0");
        return client;
    }
}

/// <summary>The chosen font isn't installed but can be downloaded: the caller may ask the user first.</summary>
internal sealed class FontNotInstalledException(MathFontInfo font)
    : Exception($"{font.Family} is not installed. {font.Note}")
{
    public MathFontInfo Font { get; } = font;
}
