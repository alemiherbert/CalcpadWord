namespace CalcpadWord;

/// <summary>How Word finds the add-in. Compiled into both the add-in and the installer, so they always agree.</summary>
internal static class AddInIdentity
{
    public const string ClassId = "587D11AB-E524-453A-80FE-CCB3017CFA8B";
    public const string ProgId = "CalcpadWord.Connect";
    public const string Name = "Calcpad for Word";
    public const string Description = "Evaluates Calcpad expressions and blocks in Word documents";
    public const string ComHostFile = "CalcpadWord.AddIn.comhost.dll";

    /// <summary>What the installer registers: it starts the bundled .NET runtime, then the COM host.</summary>
    public const string LoaderFile = "CalcpadWord.Loader.dll";
}
