using System.Reflection;

namespace CalcpadWord.Setup;

/// <summary>
/// CalcpadWord-Setup.exe [/quiet] [/uninstall]
/// Exit codes: 0 done, 1 failed, 2 cancelled or not possible here.
/// </summary>
internal static class Program
{
    public static readonly string Version =
        typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "?";

    [STAThread]
    private static int Main(string[] args)
    {
        var flags = args.Select(a => a.TrimStart('/', '-').ToLowerInvariant()).ToHashSet();
        var ui = new Ui(quiet: flags.Contains("quiet") || flags.Contains("q") || flags.Contains("silent"));
        try
        {
            if (flags.Contains("?") || flags.Contains("help") || flags.Contains("h"))
            {
                ui.Info($"{AddInIdentity.Name} {Version} setup\n\n" +
                        "CalcpadWord-Setup.exe             install for the current user\n" +
                        "CalcpadWord-Setup.exe /uninstall  remove it\n" +
                        "Add /quiet to run without questions or messages.\n\n" +
                        "Exit codes: 0 done, 1 failed, 2 cancelled.");
                return 0;
            }
            return flags.Contains("uninstall") || flags.Contains("u") ? Installer.Uninstall(ui) : Installer.Install(ui);
        }
        catch (Exception ex)
        {
            ui.Log($"FAILED: {ex}");
            ui.Error($"Setup failed: {ex.Message}\n\nDetails: {Ui.LogPath}");
            return 1;
        }
    }
}
