using System.Xml.Linq;

namespace CalcpadWord.Engine.Tests;

public class DelimiterScannerTests
{
    [Fact]
    public void Finds_expressions_in_a_sentence()
    {
        var found = DelimiterScanner.FindAll("The span is {=#val L} and {= R = q*L/2 }.");

        Assert.Equal(["#val L", "R = q*L/2"], found.Select(f => f.Source));
        Assert.Equal(12, found[0].Start);
        Assert.Equal("{=#val L}".Length, found[0].Length);
    }

    [Fact]
    public void Matches_nested_braces() =>
        Assert.Equal("$Root{x^2 - 2 @ x = 0 : 2}",
            Assert.Single(DelimiterScanner.FindAll("x = {= $Root{x^2 - 2 @ x = 0 : 2} } done")).Source);

    [Theory]
    [InlineData("{= a + b")]                // never closed
    [InlineData("{= a + b\r}")]             // closes in the next paragraph
    [InlineData("{ = a }")]                 // not the opening token
    public void Ignores_text_that_is_not_a_complete_expression(string text) =>
        Assert.Empty(DelimiterScanner.FindAll(text));

    [Fact]
    public void Skips_an_unclosed_opener_and_keeps_scanning() =>
        Assert.Equal("b", Assert.Single(DelimiterScanner.FindAll("{= a {= b}")).Source);
}

public class InlineDirectiveTests
{
    [Fact]
    public void Plain_expression_has_defaults()
    {
        var d = InlineDirective.Parse("  c = a + b ");

        Assert.Equal(new InlineDirective("c = a + b", DisplayMode.Equation, null, null, null), d);
    }

    [Fact]
    public void Reads_several_switches()
    {
        var d = InlineDirective.Parse("#val #round 3 #nosub #rad M|kNm");

        Assert.Equal(new InlineDirective("M|kNm", DisplayMode.Value, false, 3, AngleUnits.Radians), d);
    }

    [Theory]
    [InlineData("#round x/7")]
    [InlineData("#round 16 x")]
    [InlineData("#round -1 x")]
    [InlineData("#VALUE x")]
    public void Rejects_bad_switches(string source) =>
        Assert.Throws<FormatException>(() => InlineDirective.Parse(source));

    [Fact]
    public void Switches_are_case_insensitive() =>
        Assert.Equal(DisplayMode.Hidden, InlineDirective.Parse("#HIDE k = 2").Mode);
}

public class DocumentSettingsTests
{
    [Fact]
    public void Round_trips()
    {
        var s = new DocumentSettings { Decimals = 4, Substitute = false, Angles = AngleUnits.Gradians, Style = RenderStyle.PlainMath };

        Assert.Equal(s, DocumentSettings.Parse(s.Serialize()));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("garbage")]
    [InlineData("decimals=abc;angles=sideways;unknown=1")]
    public void Bad_input_gives_defaults(string? text) =>
        Assert.Equal(new DocumentSettings(), DocumentSettings.Parse(text));

    [Fact]
    public void Decimals_are_clamped() =>
        Assert.Equal(15, DocumentSettings.Parse("decimals=99").Decimals);

    [Fact]
    public void Manual_calculation_round_trips_without_changing_the_results()
    {
        var manual = new DocumentSettings { Automatic = false };

        Assert.False(DocumentSettings.Parse(manual.Serialize()).Automatic);
        Assert.True(DocumentSettings.Parse("decimals=2").Automatic);
        // Documents that calculate automatically store what they always did,
        Assert.DoesNotContain("auto", new DocumentSettings().Serialize());
        // and switching it doesn't make the results look different (no redraw).
        Assert.Equal(new DocumentSettings().ResultKey(), manual.ResultKey());
    }
}

public class ControlTagTests
{
    [Fact]
    public void Round_trips_sources_containing_colons_and_unicode()
    {
        var tag = new ControlTag(ControlState.Result, "ABCDEF0123456789", "x_1 = $Root{σ^2 - 2 @ σ = 0 : 2}");

        Assert.Equal(tag, ControlTag.Parse(tag.ToString()));
    }

    [Fact]
    public void Empty_fingerprint_and_editing_state_round_trip()
    {
        var tag = new ControlTag(ControlState.Editing, "", "a = 3m");

        Assert.Equal("calcpad:e:-:a = 3m", tag.ToString());
        Assert.Equal(tag, ControlTag.Parse(tag.ToString()));
    }

    [Theory]
    [InlineData(ControlState.Result, 'r', false)]
    [InlineData(ControlState.Editing, 'e', true)]
    [InlineData(ControlState.Error, 'x', true)]
    public void Every_state_round_trips(ControlState state, char letter, bool showsSource)
    {
        var tag = new ControlTag(state, "0123456789ABCDEF", "y = nope*2");

        Assert.Equal($"calcpad:{letter}:0123456789ABCDEF:y = nope*2", tag.ToString());
        Assert.Equal(tag, ControlTag.Parse(tag.ToString()));
        Assert.Equal(showsSource, tag.ShowsSource);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("other:r:-:x")]
    [InlineData("calcpad:r")]
    public void Foreign_or_broken_tags_are_not_calcpad(string? tag) =>
        Assert.Null(ControlTag.Parse(tag));

    [Fact]
    public void Block_tags_round_trip_with_line_breaks()
    {
        var tag = new ControlTag(ControlState.Result, "0123456789ABCDEF", "\"Title\nx = 1\n#if x > 0\n'yes\n#end if", ControlKind.Block);

        Assert.StartsWith("calcpad-block:r:", tag.ToString());
        Assert.True(ControlTag.IsCalcpad(tag.ToString()));
        Assert.Equal(tag, ControlTag.Parse(tag.ToString()));
    }

    [Fact]
    public void Existing_inline_tags_still_parse_as_inline() =>
        Assert.Equal(ControlKind.Inline, ControlTag.Parse("calcpad:r:-:a = 1")!.Kind);
}

public class BlockSourceTests
{
    [Theory]
    [InlineData(new[] { "L = 6m", "q = 2kN/m" }, false)]
    [InlineData(new[] { "#val L", "#round 3 x/7" }, false)]
    [InlineData(new[] { "L = 6m", "'The span" }, true)]
    [InlineData(new[] { "\"Heading", "x = 1" }, true)]
    [InlineData(new[] { "#if x > 0", "y = 1", "#end if" }, true)]
    [InlineData(new[] { "#round 3", "x = 1/7" }, true)]
    [InlineData(new[] { "$Plot{x @ x = 0 : 1}" }, true)]
    public void Recognises_scripts(string[] lines, bool isScript) =>
        Assert.Equal(isScript, BlockSource.LooksLikeScript(lines));

    [Theory]
    [InlineData("#if a|#else if b|#else|#end if")]
    [InlineData("#for i = 1 : 3|#while x < 1|#loop|#loop")]
    [InlineData("#repeat 3|'x|#loop")]
    [InlineData("#def m$ = 1|#def n$(x$)|x$*2|#end def")]
    [InlineData("'Line ending in a colon: is a comment:")]
    public void Accepts_well_formed_blocks(string lines) =>
        Assert.Null(BlockSource.Problem(lines.Split('|')));

    [Fact]
    public void Block_text_from_word_becomes_lines()
    {
        // Paragraphs (\r), manual line breaks (\v), a smart quote and trailing empty paragraphs.
        var source = SourceNormalizer.NormalizeBlock("\r“Title\rx = 1\v  y = 2 \r\r");

        Assert.Equal("\"Title\nx = 1\n  y = 2", source);
    }

    [Fact]
    public void A_windows_line_ending_is_one_break()
    {
        // A .cpd file checked out on Windows (CRLF): no blank line between its lines, so
        // ScriptConverter doesn't split blocks or add empty paragraphs. A blank line stays one.
        var source = SourceNormalizer.NormalizeBlock("\"Title\r\n'<hr/>\r\nx = 1\r\n\r\ny = 2\r\n");

        Assert.Equal("\"Title\n'<hr/>\nx = 1\n\ny = 2", source);
    }
}

public class BlockFenceTests
{
    [Theory]
    [InlineData("{=", true)]
    [InlineData("  {= ", true)]
    [InlineData("{= a + b }", false)]
    [InlineData("text {=", false)]
    public void Opening_fence_is_a_paragraph_of_its_own(string text, bool expected) =>
        Assert.Equal(expected, BlockFence.IsOpening(text));

    [Fact]
    public void Closing_fence_counts_open_braces()
    {
        // "$Plot{" opens a brace that the next lone "}" closes; the fence is the one after.
        Assert.Equal(2, BlockFence.FindClose(["$Plot{", "}", "}"]));
        Assert.Equal(0, BlockFence.FindClose(["}"]));
    }

    [Theory]
    [InlineData("x = 1|y = 2")]
    [InlineData("x = 1|{=|}")]
    public void Unclosed_fences_are_ignored(string following) =>
        Assert.Equal(-1, BlockFence.FindClose(following.Split('|')));
}

public class WordMarkupTests
{
    [Fact]
    public void Text_paragraphs_escape_xml_and_produce_a_valid_package()
    {
        var xml = WordMarkup.Package(WordMarkup.TextParagraph("a < b & c > \"d\"", TextLook.Error));

        var doc = XDocument.Parse(xml);
        Assert.Contains("a < b & c > \"d\"", doc.Root!.Value);
    }

    [Theory]
    [InlineData(TextLook.Source)]
    [InlineData(TextLook.Error)]
    [InlineData(TextLook.Hidden)]
    public void Source_text_is_excluded_from_spell_checking(TextLook look) =>
        Assert.Contains("<w:noProof/>", WordMarkup.TextParagraph("y = nothing*2", look));

    [Fact]
    public void Result_only_falls_back_to_the_whole_equation_when_there_is_no_result() =>
        Assert.Equal("<m:r><m:t>x</m:t></m:r>", WordMarkup.ResultOnly("<m:r><m:t>x</m:t></m:r>"));

    [Fact]
    public void Text_result_only_takes_the_last_value() =>
        Assert.Equal("5 m", WordMarkup.TextResultOnly("c = √(a^2 + b^2) = √((3 m)^2 + (4 m)^2) = 5 m"));

    [Theory]
    [InlineData("a – b", "a - b")]
    [InlineData("x = ½", "x = (1/2)")]
    [InlineData("“text”", "\"text\"")]
    [InlineData(" a\vb\r", "a b")]
    public void Normalizer_undoes_word_autocorrect(string input, string expected) =>
        Assert.Equal(expected, SourceNormalizer.Normalize(input));
}
