using Calcpad.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CalcpadWord.Engine;

/// <summary>
/// A block's output: Calcpad's HTML for it, with equation ids renumbered from 0 into
/// <see cref="Expressions"/> (the equations as Office Math, already styled).
/// </summary>
public sealed record BlockOutput(string Html, IReadOnlyList<string> Expressions)
{
    private string? _wordPackage;

    /// <summary><see cref="ToWordPackage"/>, made once: the add-in makes it on the background thread, then inserts it.</summary>
    public string WordPackage => _wordPackage ??= ToWordPackage();

    /// <summary>
    /// The output as a Flat OPC package for <c>Range.InsertXML</c>, made with Calcpad's own Word
    /// exporter (text, headings, equations, tables, plots).
    /// </summary>
    public string ToWordPackage()
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

        // Calcpad's output often ends in empty paragraphs (e.g. after a plot): drop them, they show as blank lines.
        while (body.LastChild is Paragraph last && body.ChildElements.Count > 1 && IsEmpty(last))
            last.Remove();

        // A block holds at least one paragraph, and must not end in a table.
        if (body.LastChild is null or Table)
            body.AppendChild(new Paragraph());

        doc.MainDocumentPart.Document.Save();
        return doc.ToFlatOpcString();
    }

    /// <summary>No text, picture or equation.</summary>
    private static bool IsEmpty(Paragraph p) =>
        p.InnerText.Trim().Length == 0 &&
        !p.Descendants().Any(e => e.LocalName is "drawing" or "pict" or "object" or "oMath" or "oMathPara");
}
