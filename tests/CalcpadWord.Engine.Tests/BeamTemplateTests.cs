using Xunit.Abstractions;

namespace CalcpadWord.Engine.Tests;

/// <summary>
/// The Dokimos RC beam template's script (templates/dokimos/Template source/beam, linked as Fixtures/RcBeamEc2.cpd),
/// converted and calculated as the add-in does in Word: expressions, blocks and input fields in order.
/// </summary>
public class BeamTemplateTests(ITestOutputHelper output)
{
    private static ConversionResult Conversion() =>
        ScriptConverter.Convert(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RcBeamEc2.cpd")));

    private static List<WorksheetItem> Items(ConversionResult conversion, Dictionary<string, string>? values = null)
    {
        var items = new List<WorksheetItem>();
        foreach (var part in conversion.Parts)
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
        return items;
    }

    private (List<WorksheetItem> Items, IReadOnlyList<EvaluationResult> Results) Calculate(Dictionary<string, string>? values = null)
    {
        var conversion = Conversion();
        Assert.True(conversion.Converted, conversion.KeptAsBlockReason);
        var items = Items(conversion, values);
        var results = Worksheet.Evaluate(items, new DocumentSettings());
        for (var i = 0; i < results.Count; ++i)
        {
            if (results[i].Kind == ResultKind.Error || results[i].ErrorCount > 0)
                output.WriteLine($"ERROR in {items[i].Kind} '{items[i].Source.Split('\n')[0]}': {results[i].Error}");
        }
        return (items, results);
    }

    /// <summary>The value at the end of an inline result such as "M_Ed = w_Ed·L^2/8 = … = 137.53".</summary>
    private static double Value(List<WorksheetItem> items, IReadOnlyList<EvaluationResult> results, string name)
    {
        var i = items.FindIndex(item => item.Kind == ControlKind.Inline && item.Source.StartsWith(name + " =", StringComparison.Ordinal));
        Assert.True(i >= 0, $"no expression defines {name}");
        var text = results[i].Text.Replace(" ", "").Replace("−", "-");
        return double.Parse(text[(text.LastIndexOf('=') + 1)..], System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string Summary(List<WorksheetItem> items, IReadOnlyList<EvaluationResult> results) =>
        results[items.FindLastIndex(item => item.Kind == ControlKind.Block)].Text;

    [Fact]
    public void Has_input_fields_for_everything_a_designer_sets()
    {
        var inputs = Conversion().Parts.OfType<ParagraphPart>().SelectMany(p => p.Runs).OfType<InputRun>().Select(i => i.Name);

        Assert.Equal(
            ["f_ck", "α_cc", "φ", "d_g", "f_yk", "w_max", "L", "b", "h", "c_nom", "n_t", "Φ_t", "n_c", "Φ_c", "n_w", "Φ_w", "s", "g_k,s", "q_k", "ψ_2"],
            inputs);
    }

    [Fact]
    public void The_default_beam_calculates_without_errors_and_passes()
    {
        var (items, results) = Calculate();

        Assert.DoesNotContain(results, r => r.Kind == ResultKind.Error || r.ErrorCount > 0);
        Assert.Contains("The beam passes every check.", Summary(items, results));
        Assert.DoesNotContain("FAIL", Summary(items, results));
    }

    [Fact]
    public void Values_match_a_hand_calculation()
    {
        var (items, results) = Calculate();

        // w_Ed = 1.35·(10 + 25·0.3·0.5) + 1.5·8 = 30.5625 kN/m over 6 m
        Assert.Equal(30.56, Value(items, results, "w_Ed"), 2);
        Assert.Equal(137.53, Value(items, results, "M_Ed"), 2);
        Assert.Equal(91.69, Value(items, results, "V_Ed"), 2);
        // d = 500 − 30 − 8 − 20/2
        Assert.Equal(452, Value(items, results, "d"), 2);
        Assert.Equal(17, Value(items, results, "f_cd"), 2);
        // V_Rd,c = 0.12·k·(100·ρ_l·f_ck)^(1/3)·b·d, k = 1 + √(200/452), ρ_l = 942.5/(300·452)
        Assert.Equal(74.57, Value(items, results, "V_Rd,c"), 1);
        Assert.Equal(13.27, Value(items, results, "l_d"), 2);
    }

    [Fact]
    public void A_shallow_beam_fails_and_says_where()
    {
        var (items, results) = Calculate(new() { ["h"] = "300", ["L"] = "8" });

        var summary = Summary(items, results);
        Assert.Contains("FAIL", summary);
        Assert.Contains("The beam fails at least one check", summary);
        // A design that fails is not a calculation that went wrong: no Calcpad errors.
        Assert.DoesNotContain(results, r => r.Kind == ResultKind.Error || r.ErrorCount > 0);
    }

    [Fact]
    public void Changing_an_input_changes_the_results()
    {
        var (items, before) = Calculate();
        var (_, after) = Calculate(new() { ["q_k"] = "16" });

        Assert.True(Value(items, after, "M_Ed") > Value(items, before, "M_Ed"));
    }
}
