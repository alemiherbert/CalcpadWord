using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CalcpadWord.Engine.Tests;

public class MathFontTests
{
    [Fact]
    public void The_ten_starters_are_offered_with_cambria_math_as_default()
    {
        Assert.Equal(10, MathFonts.Starters.Count);
        Assert.Equal("Cambria Math", MathFonts.Default);
        Assert.Contains(MathFonts.Starters, f => f.Family == "Latin Modern Math" && f.DownloadUrl!.EndsWith(".otf"));
        Assert.EndsWith("FiraMath-Regular.otf", MathFonts.Find("Fira Math")!.DownloadUrl);
        Assert.Null(MathFonts.Find("Minion Math"));                       // commercial: not offered
        Assert.Equal(9, MathFonts.Starters.Count(f => f.DownloadUrl is not null)); // all but Cambria Math
    }

    [Fact]
    public void The_list_is_alphabetical_with_fonts_found_on_the_computer_and_the_documents_own()
    {
        var installedMath = new HashSet<string>(["Cambria Math", "STIX Two Math", "Garamond Math"], StringComparer.OrdinalIgnoreCase);

        var choices = MathFonts.Choices(installedMath, f => f is "Cambria Math" or "Old Style Math", current: "Old Style Math");

        var labels = choices.Select(c => c.Font.Label).ToList();
        Assert.Equal(labels.Order(StringComparer.CurrentCultureIgnoreCase), labels);
        Assert.Equal(12, choices.Count);                                    // 10 starters + 1 found + the document's
        Assert.Contains(choices, c => c.Font.Family == "Garamond Math" && c.Installed && c.Font.Note == MathFonts.FoundNote);
        Assert.Contains(choices, c => c.Font.Family == "Old Style Math" && c.Installed);
        Assert.Contains(choices, c => c.Font.Family == "STIX Two Math" && c.Installed);
        Assert.Contains(choices, c => c.Font.Family == "Fira Math" && !c.Installed);
        Assert.Single(choices, c => c.Font.Family == "Cambria Math");
    }

    [Fact]
    public void Math_fonts_are_found_from_the_table_directory_alone()
    {
        var cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "cambria.ttc");
        if (!File.Exists(cambria))
            return; // Cambria isn't installed here

        using var file = File.OpenRead(cambria);
        Assert.Equal(["Cambria Math"], FontFile.MathFamilies(file));         // not Cambria, which has no MATH table

        var arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
        if (File.Exists(arial))
        {
            using var plain = File.OpenRead(arial);
            Assert.Empty(FontFile.MathFamilies(plain));
        }
        Assert.Empty(FontFile.MathFamilies(new MemoryStream("<html>not found</html>"u8.ToArray())));
        Assert.Empty(FontFile.MathFamilies(new MemoryStream([0x00, 0x01, 0x00, 0x00, 0x00, 0x09])));   // cut short
    }

    [Fact]
    public void Font_files_are_read_including_collections()
    {
        var cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "cambria.ttc");
        if (!File.Exists(cambria))
            return; // Cambria isn't installed here

        var fonts = FontFile.Read(File.ReadAllBytes(cambria));

        Assert.Contains(("Cambria Math", true), fonts);   // the math font, with its MATH table
        Assert.Contains(fonts, f => f.Family == "Cambria" && !f.IsMath);
    }

    [Fact]
    public void Something_that_is_not_a_font_reads_as_nothing() =>
        Assert.Empty(FontFile.Read("<html>not found</html>"u8));

    [Fact]
    public void Units_follow_the_chosen_math_font()
    {
        var omml = "<m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\" /></w:rPr><m:t>m</m:t></m:r>";

        Assert.Contains("w:ascii=\"Latin Modern Math\" w:hAnsi=\"Latin Modern Math\"", WordMarkup.ApplyStyle(omml, RenderStyle.Calcpad, "Latin Modern Math"));
        Assert.Contains("\"Cambria Math\"", WordMarkup.ApplyStyle(omml, RenderStyle.Calcpad));
    }

    [Theory]
    [InlineData("<m:r><m:t>F</m:t></m:r>")]                                                           // no properties
    [InlineData("<m:r><w:rPr><w:color w:val=\"0066DD\"/></w:rPr><m:t>F</m:t></m:r>")]                 // a variable
    [InlineData("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><w:rPr/><m:t>2</m:t></m:r>")]                // empty w:rPr
    [InlineData("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>=</m:t></m:r>")]                        // after m:rPr
    public void Variables_and_numbers_follow_the_chosen_math_font(string run)
    {
        var omml = $"<m:oMath>{run}<m:f><m:num>{run}</m:num><m:den>{run}</m:den></m:f></m:oMath>";

        var styled = WordMarkup.ApplyStyle(omml, RenderStyle.Calcpad, "STIX Two Math");

        Assert.Equal(3, Regex.Matches(styled, "<w:rFonts w:ascii=\"STIX Two Math\" w:hAnsi=\"STIX Two Math\"/>").Count);
        Assert.Equal(3, Regex.Matches(styled, @"<w:rPr><w:rFonts [^>]*/>(<w:color [^>]*/>)?</w:rPr><m:t>").Count); // well placed
        Assert.DoesNotContain("<m:rPr><w:rFonts", styled);
        Assert.Equal(omml, WordMarkup.ApplyStyle(omml, RenderStyle.Calcpad));       // Cambria Math: left to Word
    }

    [Fact]
    public void Real_equations_get_the_math_font_on_every_run()
    {
        var results = Worksheet.Evaluate([WorksheetItem.Inline("s = 120kN/(0.3m*0.5m)|MPa")], new DocumentSettings { MathFont = "Euler Math" });

        var runs = Regex.Matches(results[0].Omml!, @"<m:r>.*?</m:r>", RegexOptions.Singleline);
        Assert.NotEmpty(runs);
        Assert.All(runs, r => Assert.Contains("w:ascii=\"Euler Math\"", r.Value));
        Assert.DoesNotContain("Cambria", results[0].Omml);
    }

    [Fact]
    public void The_math_font_is_part_of_the_settings()
    {
        var settings = new DocumentSettings { MathFont = "TeX Gyre Pagella Math" };

        Assert.Equal(settings, DocumentSettings.Parse(settings.Serialize()));
        Assert.Equal("Cambria Math", DocumentSettings.Parse("decimals=2").MathFont);
    }

    [Fact]
    public void Changing_the_math_font_redraws_blocks_and_expressions()
    {
        WorksheetItem[] items = [WorksheetItem.Inline("a = 2m"), WorksheetItem.Block("b = a*2")];
        var cambria = Worksheet.Evaluate(items, new DocumentSettings());
        var euler = Worksheet.Evaluate(items, new DocumentSettings { MathFont = "Euler Math" });

        Assert.NotEqual(cambria[0].Fingerprint, euler[0].Fingerprint);
        Assert.NotEqual(cambria[1].Fingerprint, euler[1].Fingerprint);
        Assert.Contains("Euler Math", euler[0].Omml);
    }
}

public class ScientificNotationTests
{
    [Theory]
    [InlineData("x = 0.0000000057m")]
    [InlineData("M = 5.7nm|m")]
    public void The_times_sign_is_on_the_line_and_the_exponent_has_a_real_minus(string source)
    {
        var result = Worksheet.Evaluate([WorksheetItem.Inline(source)], new DocumentSettings())[0];
        var omml = result.Omml!;

        Assert.Contains("<m:r><m:t>×</m:t></m:r>", omml);               // an ordinary math ×
        Assert.DoesNotContain("superscript", omml);
        Assert.Matches(@"<m:sup>\s*<m:r><m:t>−</m:t></m:r>", omml);     // 10^(−9), not 10^(-9)
        Assert.DoesNotMatch(@"<m:sup>\s*<m:r>\s*<m:rPr>\s*<m:nor", omml);
    }

    [Fact]
    public void Negative_values_get_a_real_minus_and_large_exponents_are_left_alone()
    {
        var omml = Worksheet.Evaluate([WorksheetItem.Inline("y = -2*3"), WorksheetItem.Inline("z = 5.7*10^20")], new DocumentSettings());

        Assert.Contains("<m:r><m:t>−</m:t></m:r><m:r><m:t>6</m:t></m:r>", omml[0].Omml); // −6, not a hyphen
        Assert.DoesNotContain("<m:t>-</m:t>", omml[0].Omml);
        Assert.Contains("<m:r><m:t>×</m:t></m:r>", omml[1].Omml);
        Assert.DoesNotContain("−", omml[1].Omml);
    }
}

/// <summary>What the add-in corrects in Calcpad's Word equations, so that they read as Word math and the Eurocodes do.</summary>
public class EquationLookTests
{
    private static string Omml(string source, DocumentSettings? settings = null) =>
        Worksheet.Evaluate([WorksheetItem.Inline("a = 3"), WorksheetItem.Inline(source)], settings ?? new DocumentSettings())[1].Omml;

    [Theory]
    [InlineData("k = abs(-a)")]
    [InlineData("y = exp(-1)")]
    public void Minus_signs_are_math_minus_signs_not_hyphens_in_normal_text(string source)
    {
        var omml = Omml(source);

        Assert.Contains("<m:r><m:t>−</m:t></m:r>", omml);
        Assert.DoesNotContain("<m:t>-</m:t>", omml);
    }

    [Fact]
    public void Function_names_are_upright_and_regular()
    {
        var omml = Omml("y = max(sin(30); exp(a))");

        Assert.Matches(@"<m:nor/></m:rPr><w:rPr>\s*<w:rFonts [^>]*/>\s*</w:rPr><m:t>max</m:t>", omml);
        Assert.DoesNotContain("<w:b", omml);
    }

    [Fact]
    public void Subscripts_of_names_are_upright_as_in_the_Eurocodes()
    {
        var omml = Omml("A_s,req = 2*a");

        Assert.Contains("<m:sub><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>s,req</m:t></m:r></m:sub>", omml);
        Assert.Matches(@"<m:e><m:r><w:rPr>\s*<w:color [^>]*/>\s*</w:rPr><m:t>A</m:t>", omml); // the name itself stays italic
    }

    [Fact]
    public void Numbers_in_subscripts_and_the_limits_of_a_sum_are_left_alone()
    {
        var omml = Omml("n = d_2 + $Sum{i @ i = 1 : 3}", new DocumentSettings()) + Omml("d_2 = 1");

        Assert.DoesNotContain("m:sty", omml);
    }

    [Fact]
    public void Words_in_equations_are_in_the_math_font_not_the_paragraph_font()
    {
        var omml = Omml("q = if(a > 2; a; 2)");

        Assert.Matches(@"<m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=""Cambria Math"" w:hAnsi=""Cambria Math""/></w:rPr><m:t> if </m:t>", omml);
        Assert.Matches(@"<w:rFonts w:ascii=""STIX Two Math"" w:hAnsi=""STIX Two Math""/></w:rPr><m:t> else:</m:t>", Omml("q = if(a > 2; a; 2)", new DocumentSettings { MathFont = "STIX Two Math" }));
    }
}

public class CancellationTests
{
    [Fact]
    public void A_cancelled_calculation_stops()
    {
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            Worksheet.Evaluate([WorksheetItem.Inline("a = 1")], new DocumentSettings(), cancellationToken: cancel.Token));
    }

    [Fact]
    public void Cancelling_stops_a_long_calculation_quickly()
    {
        // The soil example's maps take far longer than this test allows.
        var source = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "SoilStress.cpd"));
        using var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(1));
        var timer = Stopwatch.StartNew();

        Assert.Throws<OperationCanceledException>(() =>
            Worksheet.Evaluate([WorksheetItem.Block(source)], new DocumentSettings(), cancellationToken: cancel.Token));
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5), $"took {timer.Elapsed.TotalSeconds:F1} s to stop");
    }
}
