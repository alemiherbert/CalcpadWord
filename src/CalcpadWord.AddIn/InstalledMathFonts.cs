using CalcpadWord.Engine;

namespace CalcpadWord.AddIn;

/// <summary>
/// Finds the math fonts installed on this computer: every registered font file is checked for an
/// OpenType MATH table. Only the table directories are read, and each file only once per Word
/// session, so a later look only reads fonts installed since. Any thread.
/// </summary>
internal static class InstalledMathFonts
{
    private static readonly string[] FontExtensions = [".ttf", ".otf", ".ttc", ".otc"];
    private static readonly Lock Gate = new();
    private static readonly Dictionary<string, IReadOnlyList<string>> FamiliesByFile = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The math font families installed for all users or this user.</summary>
    public static IReadOnlySet<string> Families()
    {
        var files = FontInstaller.RegisteredFiles();
        var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        lock (Gate)
        {
            var timer = System.Diagnostics.Stopwatch.StartNew();
            var read = 0;
            foreach (var file in files)
            {
                if (!FamiliesByFile.TryGetValue(file, out var found))
                {
                    FamiliesByFile[file] = found = Read(file);
                    ++read;
                }
                families.UnionWith(found);
            }
            if (read > 0)
                Log.Info($"Looked for math fonts in {read} font files ({timer.ElapsedMilliseconds} ms): {string.Join(", ", families.Order(StringComparer.OrdinalIgnoreCase))}");
        }
        return families;
    }

    private static IReadOnlyList<string> Read(string file)
    {
        if (!FontExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
            return [];
        try
        {
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.RandomAccess);
            return FontFile.MathFamilies(stream);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or EndOfStreamException)
        {
            return []; // missing, locked or damaged: not offered
        }
    }
}
