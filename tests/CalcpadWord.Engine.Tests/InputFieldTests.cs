namespace CalcpadWord.Engine.Tests;

/// <summary>Word form fields that set Calcpad variables, and Calcpad's "? {default}" inputs becoming them.</summary>
public class InputFieldTests
{
    [Theory]
    [InlineData("calcpad-input:b", "b", "")]
    [InlineData("calcpad-input:f_ck:MPa", "f_ck", "MPa")]
    [InlineData("calcpad-input:A_s,req:mm^2", "A_s,req", "mm^2")]
    [InlineData("calcpad-input:Φ_1", "Φ_1", "")]
    public void Tags_name_the_variable_and_an_optional_unit(string tag, string name, string unit)
    {
        var field = InputField.Parse(tag)!;

        Assert.Equal(new InputField(name, unit), field);
        Assert.Equal(tag, field.ToString());
    }

    [Theory]
    [InlineData("calcpad:r:-:b = 1")]
    [InlineData("calcpad-input:")]
    [InlineData("calcpad-input:2b")]
    [InlineData(null)]
    public void Other_tags_are_not_inputs(string? tag) => Assert.Null(InputField.Parse(tag));

    [Theory]
    [InlineData("", "300", "b = 300")]
    [InlineData("", " 0.85 ", "b = 0.85")]
    [InlineData("", "−12.5", "b = -12.5")] // Word's minus sign
    [InlineData("", "+3", "b = 3")]
    [InlineData("mm", "300", "b = 300mm")]
    [InlineData("mm", "0.3 m", "b = 0.3m")]
    public void Values_become_assignments(string unit, string value, string expected) =>
        Assert.Equal(expected, new InputField("b", unit).Assignment(value));

    [Theory]
    [InlineData("", "Enter a value for b.")]
    [InlineData("?", "Enter a value for b.")]
    [InlineData("0,85", "b: write the number without commas")]
    [InlineData("1,200", "b: write the number without commas")]
    [InlineData("300 mm", "b: enter a number, not \"300 mm\".")]
    [InlineData("h/2", "b: enter a number, not \"h/2\".")]
    public void Unusable_values_are_explained(string value, string message)
    {
        var ex = Assert.Throws<FormatException>(() => new InputField("b").Assignment(value));
        Assert.StartsWith(message, ex.Message);
    }

    [Fact]
    public void Input_fields_set_variables_where_they_stand()
    {
        var results = Worksheet.Evaluate(
        [
            WorksheetItem.Input(new InputField("b"), "300"),
            WorksheetItem.Input(new InputField("h"), "500"),
            WorksheetItem.Inline("A = b*h"),
            WorksheetItem.Input(new InputField("b"), "200"),
            WorksheetItem.Inline("A_2 = b*h"),
        ], new DocumentSettings());

        Assert.Equal(ResultKind.Hidden, results[0].Kind);
        Assert.Equal(ResultKind.Hidden, results[1].Kind);
        Assert.EndsWith("150000", results[2].Text.Replace(" ", ""));
        Assert.EndsWith("100000", results[4].Text.Replace(" ", ""));
    }

    [Fact]
    public void A_bad_input_fails_alone_and_says_why()
    {
        var results = Worksheet.Evaluate(
        [
            WorksheetItem.Input(new InputField("b"), "abc"),
            WorksheetItem.Inline("c = 2 + 3"),
        ], new DocumentSettings());

        Assert.Equal(ResultKind.Error, results[0].Kind);
        Assert.Equal("b: enter a number, not \"abc\".", results[0].Error);
        Assert.Equal(ResultKind.Math, results[1].Kind);
    }

    [Fact]
    public void A_changed_input_changes_the_hash_of_blocks_after_it()
    {
        string BlockHash(string b) => Worksheet.Evaluate(
            [WorksheetItem.Input(new InputField("b"), b), WorksheetItem.Block("c = b*2")], new DocumentSettings())[1].InputHash!;

        Assert.Equal(BlockHash("3"), BlockHash("3"));
        Assert.NotEqual(BlockHash("3"), BlockHash("4"));
    }

    /// <summary>Runs as text: *italic*, ~subscript~, [field name: value].</summary>
    private static string Render(ParagraphPart p) =>
        string.Concat(p.Runs.Select(r => r switch
        {
            InputRun i => $"[{i.Name}: {i.Value}{i.Unit}]",
            TextRun { Format: TextFormat.Italic } t => $"*{t.Text}*",
            TextRun { Format: TextFormat.Subscript } t => $"~{t.Text}~",
            TextRun t => t.Text,
            ExpressionRun e => "{" + e.Source + "}",
            _ => "",
        }));

    [Fact]
    public void Calcpad_inputs_become_form_fields_named_like_the_variable()
    {
        var p = Assert.IsType<ParagraphPart>(Assert.Single(ScriptConverter.Convert("'Width -'b = ? {300}'mm, height'h_w = ?{500}', length'L = ? {6}m").Parts));

        Assert.Equal("Width - *b* = [b: 300] mm, height *h*~w~ = [h_w: 500], length *L* = [L: 6m]", Render(p));
    }

    [Fact]
    public void Values_saved_after_the_script_fill_inputs_without_their_own()
    {
        var parts = ScriptConverter.Convert("a = ?\n'Text'b = ? {7}\nc = ?\n#if a > 0\nd = ?\n#end if\v1\t2\t4").Parts;

        var inputs = parts.OfType<ParagraphPart>().SelectMany(p => p.Runs).OfType<InputRun>().ToList();
        Assert.Equal([new InputRun("a", "1"), new InputRun("b", "7"), new InputRun("c", "2")], inputs);
        // Inside a block an input can't be a field: it keeps its value.
        Assert.Contains("d = 4", Assert.Single(parts.OfType<BlockPart>()).Source);
    }

    [Fact]
    public void Hidden_inputs_keep_their_value_as_a_hidden_assignment()
    {
        var p = Assert.IsType<ParagraphPart>(Assert.Single(ScriptConverter.Convert("#hide\nk = ? {1.5}\n#show").Parts));

        Assert.Equal(new ExpressionRun("#hide k = 1.5"), Assert.Single(p.Runs));
    }

    [Fact]
    public void Input_fields_are_plain_text_controls_tagged_with_the_variable()
    {
        var xml = WordMarkup.ConvertedContent([new ParagraphPart([new InputRun("f_ck", "30", "MPa")])]);

        Assert.Contains("<w:tag w:val=\"calcpad-input:f_ck:MPa\"/>", xml);
        Assert.Contains("<w:text/>", xml);
        Assert.Contains(">30</w:t>", xml);
    }
}
