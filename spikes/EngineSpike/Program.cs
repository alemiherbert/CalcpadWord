using System.Text;
using Calcpad.Core;

Console.OutputEncoding = Encoding.UTF8;
var outDir = args.Length > 0 ? args[0] : AppContext.BaseDirectory;

// ---------------------------------------------------------------------------
// 1. Inline mode: one persistent MathParser, fed one expression at a time,
//    exactly like walking a Word document top to bottom.
// ---------------------------------------------------------------------------
var settings = new MathSettings { Decimals = 3, Substitute = true, FormatEquations = true };
var p = new MathParser(settings);

string[] inline =
[
    "a = 3m",
    "b = 4m",
    "c = sqrt(a^2 + b^2)",
    "c|cm",                          // unit conversion of a result
    "f(x) = x^2 + 1",                // function definition (no result)
    "f(3)",
    "F = 10kN",
    "σ = F/(a*b)",
    "σ|MPa",
    "x_1 = 2",                       // subscripted name
    "$Root{x^2 - 2 @ x = 0 : 2}",    // solver block lives in MathParser
    "undefinedVar*2",                // error: undefined variable
    "a*",                            // error: syntax
    "a + 1m",                        // still usable after errors?
    "v = [1; 2; 3]",
    "v*a",
    "M = [1; 2|3; 4]",
    "M*M",
];

var omml = new List<(string src, string xml)>();
foreach (var expr in inline)
{
    Console.WriteLine($"IN   {expr}");
    try
    {
        p.Parse(expr);
        p.Calculate();
        Console.WriteLine($"TEXT {p}");
        Console.WriteLine($"VAL  {p.ResultAsVal}");
        Console.WriteLine($"STR  {p.ResultAsString}");
        omml.Add((expr, p.ToXml()));
    }
    catch (MathParserException ex)
    {
        Console.WriteLine($"ERR  {ex.Message}");
    }
    Console.WriteLine();
}

// One expression in all three render formats, to see what each looks like.
p.Parse("d = c/2 + a");
p.Calculate();
Console.WriteLine("=== HTML for 'd = c/2 + a'");
Console.WriteLine(p.ToHtml());
Console.WriteLine("=== OMML for 'd = c/2 + a'");
Console.WriteLine(p.ToXml());
Console.WriteLine();

// ---------------------------------------------------------------------------
// 2. Settings knobs that matter for inline output.
// ---------------------------------------------------------------------------
foreach (var (label, s) in new[]
{
    ("Substitute=false", new MathSettings { Decimals = 3, Substitute = false }),
    ("Decimals=1",       new MathSettings { Decimals = 1, Substitute = true }),
})
{
    var q = new MathParser(s);
    q.Parse("a = 3m"); q.Calculate();
    q.Parse("b = 4m"); q.Calculate();
    q.Parse("c = sqrt(a^2 + b^2)/7"); q.Calculate();
    Console.WriteLine($"[{label}] {q}");
}
Console.WriteLine();

// ---------------------------------------------------------------------------
// 3. Whole-script mode: ExpressionParser with Debug=true, to check whether each
//    rendered equation (eq-N) can be traced back to its source line.
// ---------------------------------------------------------------------------
var script = """
"Beam check
'Span and load
L = 6m
q = 12kN/m
#if L > 5m
'Long span
#end if
M = q*L^2/8
'Moment in kNm is 'M|kNm' as text
#hide
k = 2
#show
R = q*L/2
""";
var ep = new ExpressionParser { Debug = true };
ep.Settings.Math.Decimals = 2;
ep.Parse(script, calculate: true, getXml: true);
Console.WriteLine("=== ExpressionParser HTML (Debug=true)");
Console.WriteLine(ep.HtmlResult);
Console.WriteLine($"=== OpenXmlExpressions count: {ep.OpenXmlExpressions.Count}");
Console.WriteLine();

// ---------------------------------------------------------------------------
// 4. Write a Flat OPC Word package holding every inline OMML result. This is the
//    exact payload shape Range.InsertXML (COM) / insertOoxml (Office.js) accept,
//    and Word can also open the file directly to eyeball the rendering.
// ---------------------------------------------------------------------------
var body = new StringBuilder();
foreach (var (src, xml) in omml)
{
    body.Append($"<w:p><w:r><w:t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(src)}   →   </w:t></w:r>");
    body.Append($"<m:oMath>{xml}</m:oMath></w:p>");
}
var flatOpc = $"""
<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>{body}</w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
""";
var path = Path.Combine(outDir, "inline-omml-sample.xml");
File.WriteAllText(path, flatOpc, new UTF8Encoding(false));
Console.WriteLine($"Flat OPC sample written: {path}");

// Make sure the payload is well-formed XML (Word rejects InsertXML silently-ish otherwise).
System.Xml.Linq.XDocument.Parse(flatOpc);
Console.WriteLine("Flat OPC parses as well-formed XML.");
