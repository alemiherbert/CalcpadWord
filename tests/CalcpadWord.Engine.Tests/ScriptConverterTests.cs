using System.Xml.Linq;

namespace CalcpadWord.Engine.Tests;

public class ScriptConverterTests
{
    private static string Fixture(string name) =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));

    private static string Text(ParagraphPart p) =>
        string.Concat(p.Runs.Select(r => r switch
        {
            TextRun t => t.Text,
            ExpressionRun e => "{" + e.Source + "}",
            BreakRun => "|",
            _ => "",
        }));

    private static ConversionResult Convert(params string[] lines) => ScriptConverter.Convert(string.Join('\n', lines));

    [Fact]
    public void Headings_and_text_become_paragraphs_with_inline_expressions()
    {
        var parts = Convert(
            "\"Beam check",
            "'Dimensions -'a = 2m', 'b = 3m",
            "'<p><strong>Loads</strong></p>",
            "q = 5kN/m",
            "'<h4>Results</h4>").Parts;

        Assert.Collection(parts,
            p => { var h = Assert.IsType<ParagraphPart>(p); Assert.Equal(1, h.HeadingLevel); Assert.Equal("Beam check", Text(h)); },
            p => Assert.Equal("Dimensions - {a = 2m}, {b = 3m}", Text(Assert.IsType<ParagraphPart>(p))),
            p =>
            {
                var bold = Assert.IsType<ParagraphPart>(p);
                Assert.Equal("Loads", Text(bold));
                Assert.Equal(TextFormat.Bold, Assert.IsType<TextRun>(bold.Runs.Single()).Format);
            },
            p => Assert.Equal("{q = 5kN/m}", Text(Assert.IsType<ParagraphPart>(p))),
            p => { var h = Assert.IsType<ParagraphPart>(p); Assert.Equal(2, h.HeadingLevel); Assert.Equal("Results", Text(h)); });
    }

    [Fact]
    public void Settings_inside_html_comments_become_a_hidden_paragraph()
    {
        var p = Assert.IsType<ParagraphPart>(Assert.Single(Convert("'<!--'k = 1.5','PlotStep = 5'-->").Parts));

        Assert.True(p.Hidden);
        Assert.Equal("{#hide k = 1.5}{#hide PlotStep = 5}", Text(p));
    }

    [Fact]
    public void Plots_tables_and_control_flow_stay_in_blocks_and_neighbours_merge()
    {
        var parts = Convert(
            "'Plot for 'y = 0m",
            "$Plot{x^2 @ x = 0 : 1}",
            "'<table><tr><td>",
            "$Map{x*y @ x = 0 : 1 & y = 0 : 1}",
            "'</td></tr></table>",
            "'After",
            "#if y > 0",
            "'positive",
            "#end if").Parts;

        Assert.Collection(parts,
            p => Assert.Equal("Plot for {y = 0m}", Text(Assert.IsType<ParagraphPart>(p))),
            p => Assert.Equal(4, Assert.IsType<BlockPart>(p).Source.Split('\n').Length),
            p => Assert.Equal("After", Text(Assert.IsType<ParagraphPart>(p))),
            p => Assert.Equal("#if y > 0\n'positive\n#end if", Assert.IsType<BlockPart>(p).Source));
    }

    [Fact]
    public void Display_keywords_become_switches_on_the_expressions_that_follow()
    {
        var parts = Convert("#round 4", "#nosub", "'Pi is 'π", "#round default", "#hide", "k = 2", "#show", "x = k").Parts;

        Assert.Equal(["Pi is {#round 4 #nosub π}", "{#hide k = 2}", "{#nosub x = k}"],
            parts.OfType<ParagraphPart>().Select(Text));
        Assert.True(((ParagraphPart)parts[1]).Hidden);
    }

    [Fact]
    public void A_block_starts_in_the_display_state_of_its_place_in_the_script() =>
        Assert.Equal("#round 3\n$Plot{x @ x = 0 : 1}", Assert.IsType<BlockPart>(Assert.Single(Convert("#round 3", "$Plot{x @ x = 0 : 1}").Parts)).Source);

    [Fact]
    public void Calcpad_quote_rules_are_followed()
    {
        // A doubled quote is a literal quote; inside '…' a " is literal too.
        var p = Assert.IsType<ParagraphPart>(Assert.Single(Convert("'It''s \"quoted\" 'x = 1").Parts));

        Assert.Equal("It's \"quoted\" {x = 1}", Text(p));
    }

    [Theory]
    [InlineData("#md on")]
    [InlineData("#format F2")]
    [InlineData("#round n")]
    public void Scripts_using_keywords_inline_expressions_cannot_reproduce_stay_one_block(string keyword)
    {
        var result = Convert("'Text", keyword, "x = 1");

        Assert.False(result.Converted);
        Assert.Contains(keyword.Split(' ')[0], result.KeptAsBlockReason);
        Assert.IsType<BlockPart>(Assert.Single(result.Parts));
    }

    [Fact]
    public void The_soil_stress_example_becomes_word_text_with_two_blocks()
    {
        var result = ScriptConverter.Convert(Fixture("SoilStress.cpd"));

        Assert.True(result.Converted, result.KeptAsBlockReason);
        var paragraphs = result.Parts.OfType<ParagraphPart>().ToList();
        Assert.Equal("Soil stress and deformation for arbitrary surface load", Text(paragraphs[0]));
        Assert.Equal(1, paragraphs[0].HeadingLevel);
        Assert.IsType<RulePart>(result.Parts[1]);
        Assert.Contains(paragraphs, p => p.HeadingLevel == 2 && Text(p) == "Results");
        Assert.Contains(paragraphs, p => Text(p) == "Surface load - {p(x; y) = 100kPa*(1 + (x/a))}");
        Assert.Contains(paragraphs, p => p.Hidden && Text(p).Contains("{#hide PlotStep = 5}"));
        Assert.DoesNotContain(paragraphs, p => Text(p).Contains('<'));
        Assert.Equal(2, result.Parts.OfType<BlockPart>().Count()); // the first plot with the table of maps, and the last plot
    }

    [Fact]
    public void The_converted_soil_example_calculates_like_the_script()
    {
        var parts = ScriptConverter.Convert(Fixture("SoilStress.cpd")).Parts;
        var items = new List<WorksheetItem>();
        foreach (var part in parts)
        {
            if (part is BlockPart block)
                items.Add(WorksheetItem.Block(block.Source));
            else if (part is ParagraphPart paragraph)
                items.AddRange(paragraph.Runs.OfType<ExpressionRun>().Select(e => WorksheetItem.Inline(e.Source)));
        }
        items.Add(WorksheetItem.Inline("#val k*a"));

        var results = Worksheet.Evaluate(items, new DocumentSettings());

        Assert.All(results, r => Assert.True(r.Kind != ResultKind.Error && r.ErrorCount == 0, $"{r.Source}: {r.Error}"));
        Assert.Equal("3 m", results[^1].Text);
    }

    [Fact]
    public void Converted_content_is_valid_word_markup()
    {
        var body = WordMarkup.ConvertedContent(ScriptConverter.Convert(Fixture("SoilStress.cpd")).Parts);

        var xml = XDocument.Parse(WordMarkup.Package(body));
        var tags = xml.Descendants().Where(e => e.Name.LocalName == "tag").Select(e => e.Attribute(e.Name.Namespace + "val")!.Value).ToList();
        Assert.Contains(tags, t => t.StartsWith("calcpad-block:", StringComparison.Ordinal) && t.Contains("\n$Plot", StringComparison.Ordinal));
        Assert.Contains("calcpad:r:-:a = 2m", tags);
    }
}
