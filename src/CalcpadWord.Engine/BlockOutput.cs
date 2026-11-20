using Calcpad.OpenXml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace CalcpadWord.Engine;

/// <summary>
/// A block's output: Calcpad's HTML for it, with equation ids renumbered from 0 into
/// <see cref="Expressions"/> (the equations as Office Math, already styled).
/// </summary>
public sealed record BlockOutput(string Html, IReadOnlyList<string> Expressions)
{
    private string? _wordPackage;
    private double? _packageWidth;

    /// <summary><see cref="ToWordPackage"/>, made once: the add-in makes it on the background thread, then inserts it.</summary>
    public string WordPackage => WordPackageFor(null);

    /// <summary><see cref="WordPackage"/> with plots as wide as the text (in points), if known; made once per width.</summary>
    public string WordPackageFor(double? textWidth)
    {
        if (_wordPackage is null || _packageWidth != textWidth)
        {
            _wordPackage = ToWordPackage(textWidth);
            _packageWidth = textWidth;
        }
        return _wordPackage;
    }

    /// <summary>
    /// The output as a Flat OPC package for <c>Range.InsertXML</c>, made with Calcpad's own Word
    /// exporter (text, headings, equations, tables, plots). Output the exporter can't take, such
    /// as a drawing whose attributes hold an error (an input mistyped), becomes a short message
    /// in its place: one block must not stop the whole calculation.
    /// </summary>
    /// <param name="textWidth">The text width in points: plots are made that wide. Null leaves them as Calcpad drew them.</param>
    public string ToWordPackage(double? textWidth = null)
    {
        try
        {
            return Convert(textWidth);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return WordMarkup.Package(
                "<w:p><w:r><w:rPr><w:color w:val=\"C00000\"/></w:rPr><w:t xml:space=\"preserve\">" +
                System.Security.SecurityElement.Escape(
                    $"This block's output couldn't be shown in Word ({ex.Message}). Check the values it uses.") +
                "</w:t></w:r></w:p>");
        }
    }

    private string Convert(double? textWidth)
    {
        byte[] docx;
        using (var written = new MemoryStream())
        {
            new OpenXmlWriter([.. Expressions]).Convert($"<html><body>{Html}</body></html>", written);
            docx = written.ToArray();
        }

        using var stream = new MemoryStream();
        stream.Write(docx);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, isEditable: true);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // The page setup belongs to the user's document, not to this block.
        foreach (var sectionProperties in body.Elements<SectionProperties>().ToList())
            sectionProperties.Remove();

        if (FormatBorderedTables(body))
        {
            AddTableStyle(doc.MainDocumentPart);
            UseDocumentFonts(doc.MainDocumentPart);
        }
        LayOutPictures(body, textWidth);
        AddSvgFallbacks(doc.MainDocumentPart);
        KeepOnlyWhatIsUsed(doc.MainDocumentPart);

        // Calcpad's output often ends in empty paragraphs (e.g. after a plot): drop them, they show as blank lines.
        while (body.LastChild is Paragraph last && body.ChildElements.Count > 1 && IsEmpty(last))
            last.Remove();

        // A block holds at least one paragraph, and must not end in a table.
        if (body.LastChild is null or Table)
            body.AppendChild(new Paragraph());

        doc.MainDocumentPart.Document.Save();
        return doc.ToFlatOpcString();
    }

    /// <summary>
    /// Word imports every style and list definition a package brings, each time a block is drawn:
    /// Calcpad's writer adds its whole style sheet and a numbering part to every block, so a
    /// one-line PASS cost about 0.2 s to insert, and each draw left another list definition in the
    /// document. Only the styles the block uses (and those they are based on) stay, and the
    /// numbering only if it has a list.
    /// </summary>
    private static void KeepOnlyWhatIsUsed(MainDocumentPart main)
    {
        var body = main.Document!.Body!;
        if (main.NumberingDefinitionsPart is { } numbering && !body.Descendants<NumberingProperties>().Any())
            main.DeletePart(numbering);

        if (main.StyleDefinitionsPart is not { Styles: { } styles } stylesPart)
            return;

        var used = new List<string>();
        foreach (var element in body.Descendants())
        {
            var id = element switch
            {
                ParagraphStyleId p => p.Val?.Value,
                RunStyle r => r.Val?.Value,
                TableStyle t => t.Val?.Value,
                _ => null,
            };
            if (id is not null)
                used.Add(id);
        }

        var byId = styles.Elements<Style>().Where(s => s.StyleId?.Value is not null).ToDictionary(s => s.StyleId!.Value!);
        var keep = new HashSet<string>();
        foreach (var id in used)
        {
            // The style, then the one it is based on, and so on.
            for (var next = id; next is not null && keep.Add(next);)
                next = byId.TryGetValue(next, out var style) ? style.BasedOn?.Val?.Value : null;
        }

        if (keep.Count == 0)
        {
            main.DeletePart(stylesPart);
            return;
        }
        foreach (var style in byId.Values.Where(s => !keep.Contains(s.StyleId!.Value!)))
            style.Remove();
        styles.Save();
    }

    /// <summary>The title a Calcpad plot's picture gets, so the add-in can find it in the document.</summary>
    public const string PlotTitle = "Calcpad plot";

    /// <summary>
    /// A picture on a line of its own (a plot, an SVG drawing) is centred. Plots are titled
    /// <see cref="PlotTitle"/> and, given the text width, made that wide here: resizing them in
    /// Word made it lay out the pages, 1–3 s a draw on the v2 beam.
    /// </summary>
    private static void LayOutPictures(Body body, double? textWidth)
    {
        const long EmuPerPoint = 12700;
        foreach (var drawing in body.Descendants<Drawing>())
        {
            if (drawing.Inline is not { } inline)
                continue;
            // Calcpad names a plot's picture after its alt text, "Plot".
            if (inline.DocProperties is { } properties && properties.Name?.Value == "Plot")
            {
                properties.Title = PlotTitle;
                if (textWidth is double width && width > 72 && inline.Extent is { Cx.Value: > 0, Cy.Value: > 0 } extent)
                {
                    var cx = (long)(width * EmuPerPoint);
                    var cy = extent.Cy.Value * cx / extent.Cx.Value;
                    extent.Cx = cx;
                    extent.Cy = cy;
                    foreach (var size in inline.Descendants<DocumentFormat.OpenXml.Drawing.Extents>())
                    {
                        size.Cx = cx;
                        size.Cy = cy;
                    }
                }
            }

            if (drawing.Ancestors<Paragraph>().FirstOrDefault() is { } paragraph &&
                paragraph.Parent is not TableCell &&
                paragraph.InnerText.Trim().Length == 0 &&
                !paragraph.Descendants().Any(e => e.LocalName is "oMath" or "oMathPara"))
            {
                (paragraph.ParagraphProperties ??= new ParagraphProperties()).Justification =
                    new Justification { Val = JustificationValues.Center };
                // A title just above the picture stays on its page.
                if (paragraph.PreviousSibling() is Paragraph title && title.InnerText.Trim().Length > 0)
                    (title.ParagraphProperties ??= new ParagraphProperties()).KeepNext = new KeepNext();
            }
        }
    }

    /// <summary>The extension of a picture that holds its SVG, as Word writes it.</summary>
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    /// <summary>A 1×1 transparent PNG.</summary>
    private static readonly byte[] TransparentPixel = System.Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAANSURBVBhXY/j//z8DAAj8Av6IXwbgAAAAAElFTkSuQmCC");

    /// <summary>
    /// Calcpad's writer points an SVG drawing's picture straight at the SVG. Microsoft 365 makes a
    /// PNG copy when it is inserted, but Word 2019 doesn't: it shows a broken picture and saves it
    /// with no image ("The picture can't be displayed" when the document is opened again). So the
    /// picture is made as Word writes one: a PNG, with the SVG in an extension. Word 2016 and later
    /// draw the SVG; the PNG, a transparent pixel, is only for programs that can't.
    /// </summary>
    private static void AddSvgFallbacks(MainDocumentPart main)
    {
        string? fallback = null;
        foreach (var blip in main.Document!.Body!.Descendants<A.Blip>())
        {
            if (blip.Embed?.Value is not { } svg ||
                !main.TryGetPartById(svg, out var part) || part.ContentType != "image/svg+xml")
                continue;

            if (fallback is null)
            {
                var png = main.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream(TransparentPixel))
                    png.FeedData(stream);
                fallback = main.GetIdOfPart(png);
            }
            blip.Embed = fallback;
            var extensions = blip.GetFirstChild<A.BlipExtensionList>() ?? blip.AppendChild(new A.BlipExtensionList());
            if (!extensions.Elements<A.BlipExtension>().Any(e => e.Uri?.Value == SvgExtensionUri))
                extensions.AppendChild(new A.BlipExtension(new SVGBlip { Embed = svg }) { Uri = SvgExtensionUri });
        }
    }

    /// <summary>Word's built-in table style for Calcpad's bordered tables: the default the user chose.</summary>
    public const string TableStyleName = "Grid Table 1 Light";

    private const string TableStyleId = "GridTable1Light";

    /// <summary>
    /// Calcpad's bordered tables (<c>class="bordered"</c>, its table style "bcpd") as Word
    /// tables: <see cref="TableStyleName"/>, the full width of the text, text left aligned, and
    /// a first row of headings (<c>th</c>) marked as the heading row, which repeats on a new
    /// page. Calcpad's own grey borders and shading and its centred text are direct formatting
    /// that would hide the style, so they go. Returns whether there was such a table.
    /// </summary>
    private static bool FormatBorderedTables(Body body)
    {
        var any = false;
        foreach (var table in body.Descendants<Table>())
        {
            var properties = table.GetFirstChild<TableProperties>();
            if (properties?.TableStyle?.Val?.Value != "bcpd")
                continue;

            any = true;
            properties.TableStyle = new TableStyle { Val = TableStyleId };
            properties.TableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
            // Word keeps a table's margins as they were in the block: without these, none.
            properties.TableCellMarginDefault = new TableCellMarginDefault(
                "<w:tblCellMar xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:top w:w=\"28\" w:type=\"dxa\"/><w:left w:w=\"72\" w:type=\"dxa\"/>" +
                "<w:bottom w:w=\"28\" w:type=\"dxa\"/><w:right w:w=\"72\" w:type=\"dxa\"/></w:tblCellMar>");
            properties.TableLook = new TableLook
            {
                Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false,
                NoHorizontalBand = false, NoVerticalBand = true,
            };

            var rows = table.Elements<TableRow>().ToList();
            for (var r = 0; r < rows.Count; ++r)
            {
                var cells = rows[r].Elements<TableCell>().ToList();
                // Calcpad shades heading cells (th) only: a first row of them is the heading row.
                if (r == 0 && cells.Count > 0 && cells.All(c => c.TableCellProperties?.Shading is not null))
                    (rows[r].TableRowProperties ??= new TableRowProperties()).AppendChild(new TableHeader());

                foreach (var cell in cells)
                {
                    cell.TableCellProperties?.TableCellBorders?.Remove();
                    cell.TableCellProperties?.Shading?.Remove();
                    foreach (var paragraph in cell.Elements<Paragraph>())
                        paragraph.ParagraphProperties?.Justification?.Remove();
                }
            }
        }
        return any;
    }

    /// <summary>
    /// The style's definition, as Word has it built in. A document that already has the style
    /// keeps its own (Word matches inserted styles by name), so a template can restyle it.
    /// </summary>
    private static void AddTableStyle(MainDocumentPart main)
    {
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        part.Styles ??= new Styles();
        if (part.Styles.Elements<Style>().Any(s => s.StyleId?.Value == TableStyleId))
            return;

        const string border = "w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"999999\" w:themeColor=\"text1\" w:themeTint=\"66\"";
        part.Styles.AppendChild(new Style(
            "<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"table\" w:styleId=\"" + TableStyleId + "\">" +
            "<w:name w:val=\"" + TableStyleName + "\"/><w:basedOn w:val=\"TableNormal\"/><w:uiPriority w:val=\"46\"/>" +
            "<w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>" +
            "<w:tblPr><w:tblBorders>" +
            $"<w:top {border}/><w:left {border}/><w:bottom {border}/><w:right {border}/><w:insideH {border}/><w:insideV {border}/>" +
            "</w:tblBorders></w:tblPr>" +
            "<w:tblStylePr w:type=\"firstRow\"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders>" +
            "<w:bottom w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"666666\" w:themeColor=\"text1\" w:themeTint=\"99\"/>" +
            "</w:tcBorders></w:tcPr></w:tblStylePr>" +
            "<w:tblStylePr w:type=\"lastRow\"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders>" +
            "<w:top w:val=\"double\" w:sz=\"2\" w:space=\"0\" w:color=\"666666\" w:themeColor=\"text1\" w:themeTint=\"99\"/>" +
            "</w:tcBorders></w:tcPr></w:tblStylePr>" +
            "<w:tblStylePr w:type=\"firstCol\"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>" +
            "<w:tblStylePr w:type=\"lastCol\"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>" +
            "</w:style>"));
        part.Styles.Save();
    }

    /// <summary>
    /// Calcpad's package sets Calibri 12 pt as its defaults, and Word folds a package's defaults
    /// into the styles it brings in: the table style came in as Calibri 12 pt, whatever the
    /// document's font. The theme's body font and no size make the table follow the document.
    /// </summary>
    private static void UseDocumentFonts(MainDocumentPart main)
    {
        var styles = main.StyleDefinitionsPart?.Styles;
        var defaults = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (styles is null || defaults is null)
            return;

        defaults.RunFonts = new RunFonts
        {
            AsciiTheme = ThemeFontValues.MinorHighAnsi,
            HighAnsiTheme = ThemeFontValues.MinorHighAnsi,
            EastAsiaTheme = ThemeFontValues.MinorEastAsia,
            ComplexScriptTheme = ThemeFontValues.MinorBidi,
        };
        defaults.FontSize = null;
        defaults.FontSizeComplexScript = null;
        styles.Save();
    }

    /// <summary>No text, picture or equation.</summary>
    private static bool IsEmpty(Paragraph p) =>
        p.InnerText.Trim().Length == 0 &&
        !p.Descendants().Any(e => e.LocalName is "drawing" or "pict" or "object" or "oMath" or "oMathPara");
}
