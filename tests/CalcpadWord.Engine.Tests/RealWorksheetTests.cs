using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit.Abstractions;

namespace CalcpadWord.Engine.Tests;

/// <summary>Complete Calcpad worksheets, as users paste them into Word.</summary>
public class RealWorksheetTests(ITestOutputHelper output)
{
    private static string Fixture(string name) =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));

    /// <summary>
    /// Calcpad's soil stress example: double $Integral inside $Map, settings hidden in HTML
    /// comments, an HTML table assembled across lines with plots in its cells. Slow to
    /// calculate (its maps integrate at every grid point), so it also checks that an unchanged
    /// block is not calculated and drawn again.
    /// </summary>
    [Fact]
    public void Soil_stress_example_runs_as_one_block_and_is_not_redrawn_when_unchanged()
    {
        var source = SourceNormalizer.NormalizeBlock(Fixture("SoilStress.cpd"));
        Assert.True(BlockSource.LooksLikeScript(source.Split('\n')), "pasted as is, Evaluate should make it a block");

        var timer = Stopwatch.StartNew();
        var results = Worksheet.Evaluate([WorksheetItem.Block(source), WorksheetItem.Inline("#val k*a")], new DocumentSettings());
        var calculated = timer.Elapsed;
        var result = results[0];
        var package = result.Block?.ToWordPackage();
        var converted = timer.Elapsed - calculated;
        output.WriteLine($"first run: calculate {calculated.TotalSeconds:F2} s, convert to Word {converted.TotalSeconds:F2} s, package {package?.Length / 1024} KB");

        Assert.Equal(ResultKind.Block, result.Kind);
        Assert.True(result.ErrorCount == 0, $"errors: {result.Error}");
        Assert.Contains("Soil stress and deformation", result.Text);
        Assert.DoesNotContain("PlotStep", result.Text);  // settings inside <!-- --> stay hidden
        Assert.Equal("3 m", results[1].Text);            // k = 1.5 was set inside an HTML comment

        var xml = XDocument.Parse(package!);
        var pictures = Regex.Matches(package!, @"pkg:contentType=""image/").Count;
        var equations = Regex.Matches(package!, @"<m:oMath\b").Count;
        output.WriteLine($"pictures: {pictures}, equations: {equations}");
        Assert.Equal(6, pictures);                        // 4 plots + 2 maps
        Assert.True(equations > 10);
        Assert.Matches(@"<w:tbl\b", package);             // the results table
        Assert.Contains("Results", xml.Root!.Value);

        timer.Restart();
        var again = Worksheet.Evaluate(
            [WorksheetItem.Block(source, result.InputHash), WorksheetItem.Inline("#val k*a")], new DocumentSettings());
        output.WriteLine($"unchanged run: {timer.Elapsed.TotalSeconds:F2} s");

        Assert.True(again[0].Unchanged);
        Assert.Equal("3 m", again[1].Text);
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(3), $"an unchanged block took {timer.Elapsed.TotalSeconds:F1} s");
    }
}
