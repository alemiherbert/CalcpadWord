using System.Globalization;
using System.Text.RegularExpressions;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// A calculation template's Calcpad script (a fixture), converted and calculated as the add-in
/// does in Word: expressions, blocks and input fields in order.
/// </summary>
internal sealed class TemplateScript(string fixture)
{
    public ConversionResult Conversion { get; } =
        ScriptConverter.Convert(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", fixture)));

    public IEnumerable<string> InputNames =>
        Conversion.Parts.OfType<ParagraphPart>().SelectMany(p => p.Runs).OfType<InputRun>().Select(i => i.Name);

    /// <param name="values">Input values that differ from the script's defaults.</param>
    /// <param name="probes">Variables to show after the script, as inline expressions typed at its end would.</param>
    public Calculation Calculate(Dictionary<string, string>? values = null, int decimals = 2, params string[] probes)
    {
        Assert.True(Conversion.Converted, Conversion.KeptAsBlockReason);
        var items = new List<WorksheetItem>();
        foreach (var part in Conversion.Parts)
        {
            if (part is BlockPart block)
                items.Add(WorksheetItem.Block(block.Source));
            else if (part is ParagraphPart paragraph)
            {
                foreach (var run in paragraph.Runs)
                {
                    if (run is ExpressionRun e)
                        items.Add(WorksheetItem.Inline(e.Source));
                    else if (run is InputRun i)
                        items.Add(WorksheetItem.Input(new InputField(i.Name, i.Unit), values?.GetValueOrDefault(i.Name) ?? i.Value));
                }
            }
        }
        // 1*: a value the script shows in a format of its own (x = …:0.##) is shown in full here.
        items.AddRange(probes.Select(p => WorksheetItem.Inline("1*" + p)));
        return new Calculation(items, Worksheet.Evaluate(items, new DocumentSettings { Decimals = decimals }));
    }
}

internal sealed record Calculation(List<WorksheetItem> Items, IReadOnlyList<EvaluationResult> Results)
{
    public bool HasErrors => Results.Any(r => r.Kind == ResultKind.Error || r.ErrorCount > 0);

    /// <summary>Each failed item's error, for a test's failure message.</summary>
    public string Errors => string.Join("\n", Results.Where(r => r.Kind == ResultKind.Error || r.ErrorCount > 0)
        .Select(r => $"{(r.Source.Length > 60 ? r.Source[..60] + "…" : r.Source)}: {r.Error}"));

    /// <summary>The value at the end of an inline result such as "M_Ed = w_Ed·L^2/8 = … = 137.53".</summary>
    public double Value(string name)
    {
        var i = Items.FindIndex(item => item.Kind == ControlKind.Inline && item.Source.StartsWith(name + " =", StringComparison.Ordinal));
        Assert.True(i >= 0, $"no expression defines {name}");
        return LastNumber(Results[i].Text);
    }

    /// <summary>A variable calculated inside a block, read from a probe added after the script.</summary>
    public double Probed(string name)
    {
        var i = Items.FindLastIndex(item => item.Kind == ControlKind.Inline && item.Source == "1*" + name);
        Assert.True(i >= 0, $"no probe for {name}");
        return LastNumber(Results[i].Text);
    }

    /// <summary>The output text of the last block.</summary>
    public string LastBlock => Results[Items.FindLastIndex(item => item.Kind == ControlKind.Block)].Text;

    /// <summary>The output text of every block, in order.</summary>
    public string Blocks => string.Join("\n", Items.Select((item, i) => item.Kind == ControlKind.Block ? Results[i].Text : null).OfType<string>());

    private static double LastNumber(string text)
    {
        text = text.Replace(" ", "").Replace("−", "-").TrimEnd();
        var m = Regex.Match(text[(text.LastIndexOf('=') + 1)..], @"^-?\d+(\.\d+)?");
        Assert.True(m.Success, $"no number at the end of '{text}'");
        return double.Parse(m.Value, CultureInfo.InvariantCulture);
    }
}
