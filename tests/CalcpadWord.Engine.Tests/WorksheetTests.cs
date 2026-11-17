using System.Xml.Linq;

namespace CalcpadWord.Engine.Tests;

/// <summary>Inline expressions evaluated as one worksheet (same behaviour the add-in relied on before blocks).</summary>
public class InlineWorksheetTests
{
    private static readonly string[] Prelude = ["a = 3m", "b = 4m"];

    private static IReadOnlyList<EvaluationResult> Run(DocumentSettings? settings, params string[] inline) =>
        Worksheet.Evaluate([.. Prelude.Concat(inline).Select(WorksheetItem.Inline)], settings ?? new DocumentSettings());

    private static EvaluationResult Last(params string[] inline) => Run(null, inline)[^1];

    private static EvaluationResult Last(DocumentSettings settings, params string[] inline) => Run(settings, inline)[^1];

    private static void AssertWellFormed(EvaluationResult r) =>
        XDocument.Parse(WordMarkup.Package(WordMarkup.MathParagraph(r.Omml)));

    [Fact]
    public void Variables_carry_over_between_expressions()
    {
        var r = Last("c = sqrt(a^2 + b^2)");

        Assert.Equal(ResultKind.Math, r.Kind);
        Assert.EndsWith("= 5 m", r.Text);
        Assert.Contains("3 m", r.Text); // substituted values
        Assert.Contains("<m:t>c</m:t>", r.Omml);
        AssertWellFormed(r);
    }

    [Fact]
    public void Every_item_gets_its_own_result()
    {
        var results = Run(null, "c = a + b", "d = c*2");

        Assert.Equal(4, results.Count);
        Assert.All(results, r => Assert.Equal(ResultKind.Math, r.Kind));
        Assert.EndsWith("= 3 m", results[0].Text);
        Assert.EndsWith("= 14 m", results[3].Text);
    }

    [Fact]
    public void Val_switch_shows_only_the_result_with_units()
    {
        var r = Last("c = sqrt(a^2 + b^2)", "#val c|cm");

        Assert.Equal("500 cm", r.Text);
        Assert.DoesNotContain("<m:t>=</m:t>", r.Omml);
        AssertWellFormed(r);
    }

    [Fact]
    public void Val_switch_on_a_function_definition_shows_the_definition()
    {
        var r = Last("#val f(x) = x^2 + 1");

        Assert.Equal(ResultKind.Math, r.Kind);
        Assert.StartsWith("f", r.Text);
        Assert.Contains("x^2 + 1", r.Text);
    }

    [Fact]
    public void Nosub_switch_and_document_setting_both_drop_substitution()
    {
        var bySwitch = Last("#nosub c = sqrt(a^2 + b^2)");
        var bySetting = Last(new DocumentSettings { Substitute = false }, "c = sqrt(a^2 + b^2)");

        Assert.Equal(bySwitch.Text, bySetting.Text);
        Assert.DoesNotContain("3 m", bySwitch.Text);
        Assert.EndsWith("= 5 m", bySwitch.Text);
    }

    [Fact]
    public void Varsub_switch_overrides_the_document_setting() =>
        Assert.Contains("3 m", Last(new DocumentSettings { Substitute = false }, "#varsub c = a + b").Text);

    [Fact]
    public void Switches_only_affect_their_own_expression()
    {
        var withSwitches = Last("#nosub #round 4 c = sqrt(a^2 + b^2)/7", "d = c*1");
        var without = Last("c = sqrt(a^2 + b^2)/7", "d = c*1");

        Assert.Equal(without.Text, withSwitches.Text);
        Assert.EndsWith("= 0.714 m", withSwitches.Text); // Calcpad rounds small values to significant figures
    }

    [Fact]
    public void Round_switch_sets_decimals()
    {
        Assert.EndsWith("3.14", Last("π").Text);
        Assert.EndsWith("3.1416", Last("#round 4 π").Text);
    }

    [Fact]
    public void Hide_switch_calculates_without_showing()
    {
        var results = Run(null, "#hide k = 2", "k*3");

        Assert.Equal(ResultKind.Hidden, results[^2].Kind);
        Assert.EndsWith("= 6", results[^1].Text);
    }

    [Fact]
    public void Noc_switch_shows_the_expression_without_calculating()
    {
        var r = Last("#noc y = x + 1");

        Assert.Equal(ResultKind.Math, r.Kind);
        Assert.Equal("y = x + 1", r.Text);
    }

    [Theory]
    [InlineData(AngleUnits.Degrees, "sin(90)")]
    [InlineData(AngleUnits.Radians, "sin(π/2)")]
    [InlineData(AngleUnits.Gradians, "sin(100)")]
    public void Angle_setting_is_applied(AngleUnits angles, string expression) =>
        Assert.EndsWith("= 1", Last(new DocumentSettings { Angles = angles }, expression).Text);

    [Fact]
    public void Angle_switch_overrides_the_document_setting_for_one_expression()
    {
        var results = Run(null, "#rad sin(π/2)", "sin(90)");

        Assert.EndsWith("= 1", results[^2].Text);
        Assert.EndsWith("= 1", results[^1].Text);
    }

    [Fact]
    public void Errors_are_reported_and_later_expressions_still_work()
    {
        var results = Run(null, ["x = nope*2", .. Enumerable.Repeat("a*", 30), "a + b"]);

        Assert.Equal(ResultKind.Error, results[2].Kind);
        Assert.Contains("nope", results[2].Error);
        Assert.DoesNotContain("on line", results[2].Error);
        Assert.EndsWith("= 7 m", results[^1].Text);
    }

    [Fact]
    public void After_too_many_errors_the_rest_is_reported_as_not_calculated()
    {
        // Calcpad stops a worksheet after 40 errors.
        var r = Last([.. Enumerable.Repeat("nope", 45), "1 + 1"]);

        Assert.Equal(ResultKind.Error, r.Kind);
        Assert.Contains("Not calculated", r.Error);
    }

    [Theory]
    [InlineData("#bogus a", "#bogus")]
    [InlineData("#val", "no expression")]
    [InlineData("a 'comment", "Quotes")]
    [InlineData("$Plot{x @ x = 0 : 1}", "block")]
    [InlineData("a + (", "incomplete")]
    public void Inline_problems_are_reported_without_running(string source, string message)
    {
        var results = Run(null, source, "a + b");

        Assert.Equal(ResultKind.Error, results[^2].Kind);
        Assert.Contains(message, results[^2].Error, StringComparison.OrdinalIgnoreCase);
        Assert.EndsWith("= 7 m", results[^1].Text);
    }

    [Fact]
    public void Word_autocorrect_characters_are_accepted()
    {
        // Word turns " - " into an en dash and 1/2 into ½ while typing.
        var results = Run(null, "a – 1m", "a*½");

        Assert.EndsWith("= 2 m", results[^2].Text);
        Assert.EndsWith("= 1.5 m", results[^1].Text);
    }

    [Fact]
    public void Calcpad_style_keeps_variable_colour_and_plain_style_drops_it()
    {
        var calcpad = Last("c = a + b");
        var plain = Last(new DocumentSettings { Style = RenderStyle.PlainMath }, "c = a + b");

        Assert.Contains("0066DD", calcpad.Omml);
        Assert.DoesNotContain("w:color", plain.Omml);
        AssertWellFormed(plain);
    }

    [Fact]
    public void Calcpad_font_sizes_are_removed_so_equations_follow_the_paragraph() =>
        Assert.DoesNotContain("<w:sz", Last("F = 10kN/m").Omml);

    [Fact]
    public void Solver_blocks_work_inline() =>
        Assert.EndsWith("= 1.41", Last("$Root{x^2 - 2 @ x = 0 : 2}").Text);

    [Fact]
    public void Fingerprint_changes_only_when_the_output_changes()
    {
        var first = Last("c = a + b");
        var again = Last("c = a + b");
        var other = Last(new DocumentSettings { Decimals = 4 }, "c = a/b");

        Assert.Equal(first.Fingerprint, again.Fingerprint);
        Assert.NotEqual(first.Fingerprint, other.Fingerprint);
    }
}

/// <summary>Multi-line Calcpad blocks in the same worksheet as inline expressions.</summary>
public class BlockWorksheetTests
{
    private static IReadOnlyList<EvaluationResult> Run(params WorksheetItem[] items) =>
        Worksheet.Evaluate(items, new DocumentSettings());

    private static WorksheetItem Inline(string s) => WorksheetItem.Inline(s);

    private static WorksheetItem Block(params string[] lines) => WorksheetItem.Block(string.Join('\n', lines));

    [Fact]
    public void Blocks_and_inline_expressions_share_variables_both_ways()
    {
        var results = Run(
            Inline("L = 6m"),
            Block("\"Beam check", "'The span is 'L", "M = 3*L", "#if M > 10m", "'Long", "#else", "'Short", "#end if"),
            Inline("#val M"));

        var block = results[1];
        Assert.Equal(ResultKind.Block, block.Kind);
        Assert.Equal(0, block.ErrorCount);
        Assert.Contains("Beam check", block.Text);
        Assert.Contains("Long", block.Text);
        Assert.DoesNotContain("Short", block.Text);
        Assert.Equal("18 m", results[2].Text);
    }

    [Fact]
    public void Loops_repeat_their_output()
    {
        var block = Run(Block("#for i = 1 : 3", "'Value: 'i", "#loop"))[0];

        Assert.Equal(ResultKind.Block, block.Kind);
        Assert.Equal(3, block.Block!.Expressions.Count);
    }

    [Fact]
    public void Block_state_does_not_leak_into_the_rest_of_the_document()
    {
        var results = Run(
            Block("#hide", "h = 1", "#round 5", "#nosub", "#rad"),
            Inline("h + π"),
            Inline("sin(90)"));

        Assert.Equal(ResultKind.Math, results[1].Kind);
        Assert.EndsWith("= 4.14", results[1].Text);   // visible, 2 decimals
        Assert.Contains("1 + 3.14", results[1].Text); // substitution back on
        Assert.EndsWith("= 1", results[2].Text);      // degrees again
    }

    [Theory]
    [InlineData("#if 1 > 0|'yes", "#if")]
    [InlineData("#for i = 1 : 2|'x", "loop")]
    [InlineData("'a|#end if", "#end if")]
    [InlineData("#pause", "#pause")]
    [InlineData("x = max(1;", "incomplete")]   // Calcpad would join this line with the next item
    [InlineData("  |  ", "empty")]
    public void Broken_blocks_are_reported_and_the_rest_still_runs(string lines, string message)
    {
        var results = Run(Block(lines.Split('|')), Inline("1 + 1"));

        Assert.Equal(ResultKind.Error, results[0].Kind);
        Assert.Contains(message, results[0].Error, StringComparison.OrdinalIgnoreCase);
        Assert.EndsWith("= 2", results[1].Text);
    }

    [Fact]
    public void Errors_inside_a_block_are_counted_and_the_block_still_shows()
    {
        var results = Run(Block("u = nothing*2", "w = 3"), Inline("w*2"));

        Assert.Equal(ResultKind.Block, results[0].Kind);
        Assert.Equal(1, results[0].ErrorCount);
        Assert.Contains("nothing", results[0].Error);
        Assert.DoesNotContain("on line", results[0].Block!.Html);
        Assert.EndsWith("= 6", results[1].Text);
    }

    [Fact]
    public void Macros_defined_in_a_block_work_in_later_items()
    {
        var results = Run(Block("#def double$(x$) = 2*(x$)", "m = double$(3)"), Inline("double$(5)"));

        Assert.Equal(0, results[0].ErrorCount);
        Assert.EndsWith("= 10", results[1].Text);
    }

    [Fact]
    public void Macro_errors_belong_to_their_block()
    {
        var results = Run(Block("#def bad"), Inline("1 + 1"));

        Assert.Equal(ResultKind.Error, results[0].Kind);
        Assert.EndsWith("= 2", results[1].Text);
    }

    [Fact]
    public void Block_output_becomes_a_word_package_with_equations()
    {
        var block = Run(Block("\"Heading", "'Some text", "a = 2", "b = a^2"))[0];

        var package = block.Block!.ToWordPackage();

        var xml = XDocument.Parse(package);
        Assert.Contains("Heading", xml.Root!.Value);
        Assert.Contains("oMath", package);
        Assert.DoesNotContain("w:sectPr", package);
    }

    [Fact]
    public void Plots_become_pictures()
    {
        var block = Run(Block("$Plot{x^2 @ x = 0 : 2}"))[0];

        Assert.Equal(0, block.ErrorCount);
        var package = block.Block!.ToWordPackage();
        Assert.Contains("pkg:binaryData", package);
        Assert.Contains("image/", package);
    }

    [Fact]
    public void A_block_fingerprint_is_a_hash_of_its_inputs()
    {
        var first = Run(Inline("z = 2"), Block("a = z*2"))[1];
        var again = Run(Inline("z = 2"), Block("a = z*2"))[1];
        var earlierChanged = Run(Inline("z = 3"), Block("a = z*2"))[1];
        var itselfChanged = Run(Inline("z = 2"), Block("a = z*3"))[1];
        var laterChanged = Run(Inline("z = 2"), Block("a = z*2"), Inline("w = 9"))[1];

        Assert.NotNull(first.InputHash);
        Assert.Equal(first.Fingerprint, again.Fingerprint);
        Assert.Equal(first.Fingerprint, laterChanged.Fingerprint);
        Assert.NotEqual(first.Fingerprint, earlierChanged.Fingerprint);
        Assert.NotEqual(first.Fingerprint, itselfChanged.Fingerprint);
    }

    [Fact]
    public void An_unchanged_block_is_not_drawn_again_but_still_defines_its_variables()
    {
        var first = Run(Inline("z = 2"), Block("a = z*2", "$Plot{x @ x = 0 : a}"), Inline("#val a"));

        var again = Worksheet.Evaluate(
            [Inline("z = 2"), Block("a = z*2", "$Plot{x @ x = 0 : a}") with { KnownInputHash = first[1].InputHash }, Inline("#val a")],
            new DocumentSettings());

        Assert.True(again[1].Unchanged);
        Assert.Null(again[1].Block);
        Assert.Equal(first[1].Fingerprint, again[1].Fingerprint);
        Assert.Equal("4", again[2].Text);
    }

    [Fact]
    public void A_block_is_drawn_again_when_something_before_it_changes()
    {
        var first = Run(Inline("z = 2"), Block("a = z*2"));

        var again = Worksheet.Evaluate(
            [Inline("z = 5"), Block("a = z*2") with { KnownInputHash = first[1].InputHash }, Inline("#val a")],
            new DocumentSettings());

        Assert.False(again[1].Unchanged);
        Assert.NotNull(again[1].Block);
        Assert.Equal("10", again[2].Text);
    }

    [Theory]
    [InlineData("r = random(1)")]
    [InlineData("#read M from data.csv")]
    public void Blocks_that_read_files_or_random_numbers_are_always_drawn_again(string line)
    {
        var first = Run(Block(line, "a = 1"), Block("b = 2"));

        var again = Worksheet.Evaluate(
            [Block(line, "a = 1") with { KnownInputHash = first[0].InputHash }, Block("b = 2") with { KnownInputHash = first[1].InputHash }],
            new DocumentSettings());

        Assert.False(again[0].Unchanged);
        Assert.False(again[1].Unchanged); // later blocks may depend on what was read
    }
}
