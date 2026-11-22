namespace CalcpadWord.Engine.Tests;

/// <summary>The Calcpad tab's symbol galleries, and the operators typed with two keys.</summary>
public class SymbolTests
{
    private static IReadOnlyList<EvaluationResult> Run(params string[] inline) =>
        Worksheet.Evaluate([.. inline.Select(WorksheetItem.Inline)], new DocumentSettings());

    private static IEnumerable<Symbol> Kind(SymbolKind kind) => SymbolPalette.All.Where(s => s.Kind == kind);

    [Theory]
    [InlineData("a >= b", "a ≥ b")]
    [InlineData("a<=b", "a≤b")]
    [InlineData("#if a == b && c != d || e >= f", "#if a ≡ b ∧ c ≠ d ∨ e ≥ f")]
    [InlineData("a === b", "a ≡ b")]
    [InlineData("a !== b", "a ≠ b")]
    [InlineData("a %% b + c // d", "a ⦼ b + c ÷ d")]
    [InlineData("a ^^ b", "a ⊕ b")]
    [InlineData("x <* 2", "x ← 2")]
    [InlineData("z = 2<<45°", "z = 2∠45°")]
    [InlineData("a < b > c = d", "a < b > c = d")]        // single keys stay
    [InlineData("a > = b", "a > = b")]                    // only when typed together
    public void Two_key_operators_become_calcpad_symbols(string typed, string expected) =>
        Assert.Equal(expected, SourceNormalizer.Normalize(typed));

    [Theory]
    [InlineData("'a >= b, c == d")]                                    // a comment, to the end of the line
    [InlineData("\"Title: a <= b")]                                    // a title
    [InlineData("'<p class=\"ok\">PASS - a >= b</p>")]                 // HTML in a comment, with " inside '
    [InlineData("#include C:\\calcs\\a<<b.cpd")]                       // a file name
    [InlineData("#read M from https://example.com/m.csv")]
    public void Comments_and_file_names_are_left_alone(string line) =>
        Assert.Equal(line, SourceNormalizer.Normalize(line));

    [Fact]
    public void Math_after_a_closed_comment_is_replaced() =>
        Assert.Equal("'if' a ≥ b 'then' c ≡ d", SourceNormalizer.Normalize("'if' a >= b 'then' c == d"));

    [Fact]
    public void Block_lines_are_replaced_one_by_one() =>
        Assert.Equal("#if a ≥ b\n'a >= b\nc = 1\n#end if",
            SourceNormalizer.NormalizeBlock("#if a >= b\r'a >= b\rc = 1\r#end if"));

    [Fact]
    public void Typed_operators_calculate_like_the_symbols()
    {
        var typed = Run("a = 3", "b = 2", SourceNormalizer.Normalize("(a >= b) + (a == b) + (a != b) + (a <= b)"),
            SourceNormalizer.Normalize("(a > 2 && b > 2) + (a > 2 || b > 2)"), SourceNormalizer.Normalize("7 %% 4"));

        Assert.All(typed, r => Assert.Equal(ResultKind.Math, r.Kind));
        Assert.EndsWith("= 2", typed[2].Text);   // 1 + 0 + 1 + 0
        Assert.EndsWith("= 1", typed[3].Text);   // 0 + 1
        Assert.EndsWith("= 3", typed[4].Text);
    }

    [Fact]
    public void Every_letter_and_mark_is_allowed_in_a_name()
    {
        var names = Kind(SymbolKind.Letter).Concat(Kind(SymbolKind.Mark))
            .Select(s => char.IsLetter(s.Character) ? $"{s.Character}_1" : $"x{s.Character}").ToList();
        var results = Run([.. names.Select(n => $"{n} = 2"), string.Join(" + ", names)]);

        Assert.All(results, r => Assert.True(r.Kind == ResultKind.Math, $"{r.Source}: {r.Error}"));
        Assert.EndsWith($"= {2 * names.Count}", results[^1].Text);
    }

    [Fact]
    public void Every_operator_is_one_calcpad_knows()
    {
        // ← needs a block and ∠ complex numbers; the rest work between two numbers.
        var operators = Kind(SymbolKind.Operator).Select(s => s.Character).Where(c => c is not ('←' or '∠')).ToList();
        var results = Run([.. operators.Select(op => $"7 {op} 2")]);

        Assert.All(results, r => Assert.True(r.Kind == ResultKind.Math, $"{r.Source}: {r.Error}"));
    }

    [Fact]
    public void Each_operator_shows_the_keys_that_type_it()
    {
        foreach (var symbol in Kind(SymbolKind.Operator))
        {
            Assert.NotNull(symbol.Typed);
            Assert.Equal(symbol.Character.ToString(), SourceNormalizer.ReplaceOperators(symbol.Typed));
        }
    }

    [Fact]
    public void The_ribbon_with_its_galleries_is_valid_custom_ui()
    {
        // Word silently drops a whole tab whose markup breaks the schema, so check it here.
        var xml = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Ribbon.xml"))
            .Replace("<!--SYMBOL_GALLERIES-->", SymbolPalette.RibbonXml())
            .Replace("<!--CONTEXT_MENUS-->", "");
        using var stream = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        doc.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());
        var part = doc.AddRibbonAndBackstageCustomizationsPart();
        using (var writer = new StreamWriter(part.GetStream(FileMode.Create)))
            writer.Write(xml);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010)
            .Validate(doc).Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors));
        Assert.NotNull(part.CustomUI); // parsed as customUI 2009/07, not ignored
        var gallery = Assert.Single(part.CustomUI.Descendants<DocumentFormat.OpenXml.Office2010.CustomUI.Gallery>());
        var items = gallery.Elements<DocumentFormat.OpenXml.Office2010.CustomUI.Item>().ToList();
        Assert.Equal(SymbolPalette.All.Select(s => SymbolPalette.ItemId(s.Character)), items.Select(i => i.Id!.Value));
        Assert.Equal(SymbolPalette.VisibleRows * SymbolPalette.Columns,
            part.CustomUI.Descendants<DocumentFormat.OpenXml.Office2010.CustomUI.Button>().Count(b => SymbolPalette.FromSlotId(b.Id) is not null));
    }

    [Fact]
    public void The_validator_catches_bad_markup()
    {
        using var stream = new MemoryStream();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        doc.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());
        var part = doc.AddRibbonAndBackstageCustomizationsPart();
        using (var writer = new StreamWriter(part.GetStream(FileMode.Create)))
            writer.Write("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"t\" label=\"T\">" +
                         "<group id=\"g\" label=\"G\"><gallery id=\"x\" label=\"X\" columns=\"0\" bogus=\"1\" /></group></tab></tabs></ribbon></customUI>");

        Assert.NotEmpty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010).Validate(doc));
    }

    [Fact]
    public void The_grid_is_two_rows_of_ten_operators_first()
    {
        Assert.All(SymbolPalette.Rows, row => Assert.InRange(row.Count, 1, SymbolPalette.Columns));
        Assert.Equal((10, 2), (SymbolPalette.Columns, SymbolPalette.VisibleRows));   // as the user asked: 2 × 10
        // The first view: the operators for #if and if(), then ∠ and the units and marks.
        Assert.All(SymbolPalette.Rows[0], s => Assert.Equal(SymbolKind.Operator, s.Kind));
        Assert.Equal("∠°′″‴‾øØ∡‰", string.Concat(SymbolPalette.Rows[1].Select(s => s.Character)));
        Assert.Equal('α', SymbolPalette.At(2, 0)!.Character);                        // Greek starts a row
        Assert.Null(SymbolPalette.At(SymbolPalette.Rows.Count, 0));
        Assert.Equal(SymbolPalette.All.Count, SymbolPalette.All.Select(s => s.Character).Distinct().Count());
    }

    [Fact]
    public void Slot_ids_give_their_place_back()
    {
        var ids = SymbolPalette.SlotIds().ToList();

        Assert.Equal(SymbolPalette.VisibleRows * SymbolPalette.Columns, ids.Distinct().Count());
        Assert.Equal((1, 9), SymbolPalette.FromSlotId(SymbolPalette.SlotId(1, 9)));
        Assert.Null(SymbolPalette.FromSlotId("CalcpadSlot_2_0"));
        Assert.Null(SymbolPalette.FromSlotId("CalcpadSlot_x"));
    }

    [Fact]
    public void Complex_numbers_are_a_document_setting()
    {
        var complex = new DocumentSettings { Complex = true };

        Assert.True(DocumentSettings.Parse(complex.Serialize()).Complex);
        Assert.DoesNotContain("complex", new DocumentSettings().Serialize());   // documents without it read as before
        Assert.NotEqual(new DocumentSettings().ResultKey(), complex.ResultKey()); // switching it recalculates
    }

    [Fact]
    public void Complex_mode_calculates_with_i_and_phasors()
    {
        IReadOnlyList<EvaluationResult> Run(bool complex, params string[] inline) =>
            Worksheet.Evaluate([.. inline.Select(WorksheetItem.Inline)], new DocumentSettings { Complex = complex });

        var real = Run(false, "z = sqrt(-4)");
        var complexResults = Run(true, "z = sqrt(-4)", "w = (1 + 2i)*(1 - 2i)", SourceNormalizer.Normalize("p = 2<<90°"));

        Assert.DoesNotContain("2i", real[0].Text);
        Assert.All(complexResults, r => Assert.True(r.Kind == ResultKind.Math, $"{r.Source}: {r.Error}"));
        Assert.EndsWith("2i", complexResults[0].Text);
        Assert.EndsWith("= 5", complexResults[1].Text);
        Assert.Contains("2i", complexResults[2].Text);   // 2∠90° = 2i
    }

    [Fact]
    public void Item_ids_are_unique_and_give_the_character_back()
    {
        var all = SymbolPalette.All.Select(s => s.Character).ToList();

        Assert.Equal(all.Count, all.Distinct().Count());
        Assert.All(all, c => Assert.Equal(c, SymbolPalette.FromItemId(SymbolPalette.ItemId(c))));
        Assert.Null(SymbolPalette.FromItemId("CalcpadSym_"));
        Assert.Null(SymbolPalette.FromItemId("Recalculate"));
    }
}
