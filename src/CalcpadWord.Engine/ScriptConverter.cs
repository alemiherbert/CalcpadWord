using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace CalcpadWord.Engine;

[Flags]
public enum TextFormat { None = 0, Bold = 1, Italic = 2, Underline = 4, Subscript = 8, Superscript = 16 }

public abstract record DocumentPart;

/// <param name="HeadingLevel">0 for body text, 1–4 for Word's Heading 1–4.</param>
/// <param name="Hidden">Nothing visible (settings in an HTML comment, or text in a #hide section): hidden text, paragraph mark included.</param>
public sealed record ParagraphPart(IReadOnlyList<RunPart> Runs, int HeadingLevel = 0, bool Hidden = false) : DocumentPart;

/// <summary>A horizontal rule (Calcpad's &lt;hr/&gt;).</summary>
public sealed record RulePart : DocumentPart;

/// <summary>Lines that need Calcpad to draw them: plots, maps, HTML tables, #if, loops, macros.</summary>
public sealed record BlockPart(string Source) : DocumentPart;

public abstract record RunPart;

public sealed record TextRun(string Text, TextFormat Format = TextFormat.None) : RunPart;

/// <summary>An inline Calcpad expression, switches included (e.g. "#hide k = 1.5").</summary>
public sealed record ExpressionRun(string Source) : RunPart;

public sealed record BreakRun : RunPart;

/// <summary>
/// Calcpad's input field (<c>b = ? {300}</c>): a Word form field holding the value, which sets the
/// variable where it stands (see <see cref="InputField"/>). The variable name is written before it.
/// </summary>
public sealed record InputRun(string Name, string Value, string Unit = "") : RunPart;

/// <param name="KeptAsBlockReason">Set when the script can't be converted and stays one block.</param>
public sealed record ConversionResult(IReadOnlyList<DocumentPart> Parts, string? KeptAsBlockReason = null)
{
    public bool Converted => KeptAsBlockReason is null;
}

/// <summary>
/// Turns a Calcpad script into Word content: headings and comment text become native Word
/// paragraphs with the expressions in them as inline Calcpad expressions; only what needs
/// Calcpad to draw it (plots, maps, HTML tables and lists, #if, loops, macros) stays in blocks.
/// Display keywords (#round, #rad, #nosub, #val, #hide, …) become switches on the inline
/// expressions that follow them. Evaluated in document order, the result computes exactly
/// like the script.
/// </summary>
public static partial class ScriptConverter
{
    private const char PlaceholderStart = '';
    private const char PlaceholderEnd = '';

    // HTML that Word paragraphs can't express: such lines (to the closing tag) stay in a block.
    private const string StructuralTags = "table|thead|tbody|tfoot|tr|td|th|caption|colgroup|col|div|ul|ol|li|dl|dt|dd|img|svg|canvas|iframe|figure|figcaption|object|embed|form|input|select|textarea|button|script|style|pre|blockquote|details|summary|section|article|aside|nav|header|footer|video|audio";

    [GeneratedRegex(@"<\s*(?<close>/)?\s*(?<tag>" + StructuralTags + @")\b[^>]*?(?<self>/)?\s*>", RegexOptions.IgnoreCase)]
    private static partial Regex StructuralTag();

    [GeneratedRegex("(\\d+)")]
    private static partial Regex Placeholder();

    // Void elements never close: they count as open-and-closed.
    private static readonly HashSet<string> VoidTags = new(StringComparer.OrdinalIgnoreCase) { "img", "col", "input", "embed" };

    // Calcpad's input field: "b = ?", "b = ? {300}", "L = ? {6}m".
    [GeneratedRegex(@"^(?<name>[\p{L}_][\p{L}\p{N}_′″‴⁗,.]*)\s*=\s*\?\s*(\{(?<value>[^{}]*)\})?\s*(?<unit>[\p{L}°][\p{L}\p{N}°/·*^.]*)?$")]
    private static partial Regex InputExpression();

    // The same inside a longer line (blocks), and a value as Calcpad stores it after the script.
    [GeneratedRegex(@"(?<name>[\p{L}_][\p{L}\p{N}_′″‴⁗,.]*)\s*=\s*\?\s*(\{(?<value>[^{}]*)\})?")]
    private static partial Regex InputInLine();

    [GeneratedRegex(@"^-?(\d+\.?\d*|\.\d+)$")]
    private static partial Regex FormValue();

    public static ConversionResult Convert(string script)
    {
        var (body, formValues) = SplitFormData(script);
        var source = SourceNormalizer.NormalizeBlock(body);
        var lines = source.Length == 0 ? [] : source.Split('\n');
        if (Unsupported(lines) is { } reason)
            return new ConversionResult([new BlockPart(source)], reason);

        var parts = new List<DocumentPart>();
        var state = new SwitchState();
        var block = new List<string>();
        int flowDepth = 0, macroDepth = 0, htmlDepth = 0;

        foreach (var raw in lines)
        {
            var line = raw.Trim();
            if (flowDepth > 0 || macroDepth > 0 || htmlDepth > 0)
            {
                AddToBlock(raw);
                Track(line);
                continue;
            }
            if (line.Length == 0)
            {
                FlushBlock();
                if (!state.Hidden)
                    parts.Add(new ParagraphPart([]));
                continue;
            }
            if (line[0] == '#')
            {
                if (state.Apply(line))
                    continue; // a display switch: applied to the inline expressions that follow

                AddToBlock(raw); // #if, loops, macros, #include, #read, #const, ...
                Track(line);
                continue;
            }
            if (line.StartsWith("$plot", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("$map", StringComparison.OrdinalIgnoreCase))
            {
                AddToBlock(raw);
                continue;
            }

            var tokens = Tokenize(line);
            var (markup, expressions) = Markup(tokens);
            if (StructuralTag().IsMatch(markup))
            {
                AddToBlock(raw);
                htmlDepth = Math.Max(0, HtmlDepthChange(markup));
                continue;
            }

            FlushBlock();
            var isHeading = tokens.Count > 0 && tokens[0].Kind == TokenKind.Heading;
            parts.AddRange(ConvertLine(markup, expressions, isHeading, state, formValues));
        }
        FlushBlock();
        return new ConversionResult(parts);

        void AddToBlock(string line)
        {
            if (block.Count == 0)
                block.AddRange(state.KeywordLines()); // carry the active switches into the block
            block.Add(FixedInputs(line, formValues));
        }

        void FlushBlock()
        {
            if (block.Count == 0)
                return;
            parts.Add(new BlockPart(string.Join('\n', block)));
            block.Clear();
        }

        void Track(string line)
        {
            var lower = line.ToLowerInvariant();
            if (IsKeyword(lower, "#end if")) --flowDepth;
            else if (IsKeyword(lower, "#if")) ++flowDepth;
            else if (IsKeyword(lower, "#loop")) --flowDepth;
            else if (IsKeyword(lower, "#for") || IsKeyword(lower, "#while") || IsKeyword(lower, "#repeat")) ++flowDepth;
            else if (IsKeyword(lower, "#end def")) --macroDepth;
            else if (IsKeyword(lower, "#def") && !line.Contains('=')) ++macroDepth;
            else if (!lower.StartsWith('#'))
                htmlDepth = Math.Max(0, htmlDepth + HtmlDepthChange(line));
        }
    }

    /// <summary>
    /// Calcpad saves the values of input fields after the script: a '\v', then the values
    /// separated by tabs. They count only if there is one for every input field without a {value}.
    /// </summary>
    private static (string Script, Queue<string> Values) SplitFormData(string script)
    {
        var at = script.LastIndexOf('\v');
        if (at >= 0)
        {
            var data = script[(at + 1)..].Trim();
            var values = data.Split('\t').Select(v => v.Trim()).ToArray();
            var fields = InputInLine().Matches(script[..at]).Count(m => !m.Groups["value"].Success);
            if (data.Length > 0 && data.IndexOfAny(['\r', '\n']) < 0 && values.All(v => FormValue().IsMatch(v)) &&
                values.Length == fields)
            {
                return (script[..at], new Queue<string>(values));
            }
        }
        return (script, new Queue<string>());
    }

    /// <summary>The value of the next input field: its own {default}, else the next saved value.</summary>
    private static string InputValue(Group braces, Queue<string> formValues) =>
        braces.Success ? braces.Value.Trim() : formValues.TryDequeue(out var saved) ? saved : string.Empty;

    /// <summary>Inside a block an input can't be a form field: it keeps its value, as an ordinary assignment.</summary>
    private static string FixedInputs(string line, Queue<string> formValues) =>
        InputInLine().Replace(line, m =>
        {
            var value = InputValue(m.Groups["value"], formValues);
            return value.Length == 0 ? m.Value : $"{m.Groups["name"].Value} = {value}";
        });

    /// <summary>A variable name as Calcpad writes it: italic, with the part after the first _ as a subscript.</summary>
    private static IEnumerable<RunPart> NameRuns(string name)
    {
        var sub = name.IndexOf('_');
        if (sub <= 0)
        {
            yield return new TextRun(name, TextFormat.Italic);
            yield break;
        }
        yield return new TextRun(name[..sub], TextFormat.Italic);
        yield return new TextRun(name[(sub + 1)..], TextFormat.Subscript);
    }

    /// <summary>Keywords whose effect inline expressions can't reproduce: the script then stays one block.</summary>
    private static string? Unsupported(IEnumerable<string> lines)
    {
        foreach (var raw in lines)
        {
            var lower = raw.Trim().ToLowerInvariant();
            foreach (var keyword in new[] { "#md", "#format", "#split", "#phasor", "#novar", "#input", "#pause" })
                if (IsKeyword(lower, keyword))
                    return $"it uses {keyword}";
            if (IsKeyword(lower, "#round") && lower.Length > 6 && lower[6..].Trim() is var n && n != "default" &&
                !int.TryParse(n, NumberStyles.None, CultureInfo.InvariantCulture, out _))
                return "it uses #round with a calculated value";
        }
        return null;
    }

    // ---- One line of text ---------------------------------------------------------------

    private static IEnumerable<DocumentPart> ConvertLine(
        string markup, IReadOnlyList<string> expressions, bool isHeading, SwitchState state, Queue<string> formValues)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(markup);
        var output = new List<DocumentPart>();
        var runs = new List<RunPart>();
        var level = isHeading ? 1 : 0;

        Walk(doc.DocumentNode, TextFormat.None);
        Flush();
        return output;

        void Walk(HtmlNode node, TextFormat format)
        {
            switch (node.NodeType)
            {
                case HtmlNodeType.Comment:
                    // <!-- 'x = 1' --> : Calcpad calculates the expressions but shows nothing.
                    foreach (Match m in Placeholder().Matches(node.InnerHtml))
                        AddExpression(expressions[int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)], hidden: true);
                    return;

                case HtmlNodeType.Text:
                    AddText(HtmlEntity.DeEntitize(node.InnerText), format);
                    return;
            }

            var name = node.Name.ToLowerInvariant();
            switch (name)
            {
                case "h1" or "h2" or "h3" or "h4" or "h5" or "h6":
                    Flush();
                    level = Math.Max(1, name[1] - '0' - 2); // Calcpad titles are h3: h1–h3 → Heading 1, h4 → 2, …
                    WalkChildren(node, format);
                    Flush();
                    return;
                case "p":
                    Flush();
                    WalkChildren(node, format);
                    Flush();
                    return;
                case "hr":
                    Flush();
                    output.Add(new RulePart());
                    return;
                case "br":
                    runs.Add(new BreakRun());
                    return;
                case "b" or "strong":
                    WalkChildren(node, format | TextFormat.Bold);
                    return;
                case "i" or "em":
                    WalkChildren(node, format | TextFormat.Italic);
                    return;
                case "u" or "ins":
                    WalkChildren(node, format | TextFormat.Underline);
                    return;
                case "sub":
                    WalkChildren(node, format | TextFormat.Subscript);
                    return;
                case "sup":
                    WalkChildren(node, format | TextFormat.Superscript);
                    return;
                default:
                    WalkChildren(node, format); // span, font, a, small, …: keep the text
                    return;
            }
        }

        void WalkChildren(HtmlNode node, TextFormat format)
        {
            foreach (var child in node.ChildNodes)
                Walk(child, format);
        }

        void AddText(string text, TextFormat format)
        {
            var at = 0;
            foreach (Match m in Placeholder().Matches(text))
            {
                if (m.Index > at)
                    runs.Add(new TextRun(text[at..m.Index], format));
                AddExpression(expressions[int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)], hidden: false);
                at = m.Index + m.Length;
            }
            if (at < text.Length)
                runs.Add(new TextRun(text[at..], format));
        }

        void AddExpression(string expression, bool hidden)
        {
            if (InputExpression().Match(expression) is not { Success: true } input)
            {
                runs.Add(new ExpressionRun(hidden ? "#hide " + expression : state.Prefix(expression)));
                return;
            }

            var (name, unit) = (input.Groups["name"].Value, input.Groups["unit"].Value);
            var value = InputValue(input.Groups["value"], formValues);
            if (hidden || state.Hidden)
            {
                // Nobody would see the field: the value stays fixed.
                runs.Add(new ExpressionRun($"#hide {name} = {(value.Length == 0 ? "?" : value + unit)}"));
                return;
            }
            if (runs.LastOrDefault() is TextRun { Text.Length: > 0 } before && !char.IsWhiteSpace(before.Text[^1]) && before.Text[^1] is not ('(' or '['))
                runs[^1] = before with { Text = before.Text + " " };
            runs.AddRange(NameRuns(name));
            runs.Add(new TextRun(" = "));
            runs.Add(new InputRun(name, value, unit));
        }

        void Flush()
        {
            var spaced = Spaced(runs);
            runs.Clear();
            if (spaced.Count > 0)
            {
                var visible = spaced.Any(r => r is TextRun t ? t.Text.Trim().Length > 0 : r is InputRun || r is ExpressionRun e && !e.Source.StartsWith("#hide", StringComparison.Ordinal));
                var hidden = state.Hidden || !visible;
                output.Add(new ParagraphPart(spaced, hidden ? 0 : level, hidden));
            }
            level = isHeading ? 1 : 0;
        }
    }

    /// <summary>Trims the paragraph and puts spaces between text and expressions where Word needs them.</summary>
    private static List<RunPart> Spaced(List<RunPart> runs)
    {
        var result = new List<RunPart>();
        foreach (var run in runs)
        {
            if (run is TextRun { Text.Length: 0 })
                continue;

            var previous = result.LastOrDefault();
            if (IsShown(run))
            {
                if (previous is TextRun pt && pt.Text.Length > 0 && !char.IsWhiteSpace(pt.Text[^1]) && pt.Text[^1] is not ('(' or '['))
                    result[^1] = pt with { Text = pt.Text + " " };
                else if (previous is not null && IsShown(previous))
                    result.Add(new TextRun(" "));
            }
            else if (run is TextRun t && previous is not null && IsShown(previous) &&
                     !char.IsWhiteSpace(t.Text[0]) && t.Text[0] is not (',' or '.' or ';' or ':' or ')' or ']' or '!' or '?'))
            {
                result.Add(new TextRun(" ", t.Format));
            }
            result.Add(run);
        }

        // Collapse runs of spaces and trim the ends.
        for (var i = 0; i < result.Count; ++i)
            if (result[i] is TextRun t)
                result[i] = t with { Text = Regex.Replace(t.Text, @"\s+", " ") };
        if (result.FirstOrDefault() is TextRun first)
            result[0] = first with { Text = first.Text.TrimStart() };
        if (result.LastOrDefault() is TextRun last)
            result[^1] = last with { Text = last.Text.TrimEnd() };
        result.RemoveAll(r => r is TextRun { Text.Length: 0 });
        return result;

        // A visible expression or an input field: text around it is spaced off.
        static bool IsShown(RunPart r) => r is InputRun || r is ExpressionRun e && !e.Source.StartsWith("#hide", StringComparison.Ordinal);
    }

    // ---- Calcpad's tokenizer --------------------------------------------------------------

    private enum TokenKind { Expression, Text, Heading }

    private readonly record struct Token(TokenKind Kind, string Value);

    /// <summary>
    /// As Calcpad splits a line: ' and " open a comment that the same character closes (the
    /// other one is literal inside it; a doubled one is an escaped quote). The rest are expressions.
    /// </summary>
    private static List<Token> Tokenize(string line)
    {
        var tokens = new List<Token>();
        var current = new StringBuilder();
        var separator = ' ';
        for (var i = 0; i < line.Length; ++i)
        {
            var c = line[i];
            if (c is '\'' or '"' && (separator == ' ' || separator == c))
            {
                if (separator == c && i + 1 < line.Length && line[i + 1] == c)
                {
                    current.Append(c);
                    ++i;
                    continue;
                }
                Add();
                separator = separator == c ? ' ' : c;
            }
            else
            {
                current.Append(c);
            }
        }
        Add();
        return tokens;

        void Add()
        {
            var value = current.ToString();
            current.Clear();
            var kind = separator switch { '"' => TokenKind.Heading, '\'' => TokenKind.Text, _ => TokenKind.Expression };
            if (kind == TokenKind.Expression ? value.Trim().Length > 0 : value.Length > 0)
                tokens.Add(new Token(kind, kind == TokenKind.Expression ? value.Trim() : value));
        }
    }

    /// <summary>The line's text as HTML with each expression replaced by a placeholder.</summary>
    private static (string Markup, List<string> Expressions) Markup(List<Token> tokens)
    {
        var sb = new StringBuilder();
        var expressions = new List<string>();
        foreach (var token in tokens)
        {
            if (token.Kind == TokenKind.Expression)
            {
                sb.Append(PlaceholderStart).Append(expressions.Count.ToString(CultureInfo.InvariantCulture)).Append(PlaceholderEnd);
                expressions.Add(token.Value);
            }
            else
            {
                sb.Append(token.Value);
            }
        }
        return (sb.ToString(), expressions);
    }

    private static int HtmlDepthChange(string text)
    {
        var depth = 0;
        foreach (Match m in StructuralTag().Matches(text))
        {
            if (m.Groups["self"].Success || VoidTags.Contains(m.Groups["tag"].Value))
                continue;
            depth += m.Groups["close"].Success ? -1 : 1;
        }
        return depth;
    }

    private static bool IsKeyword(string lowerLine, string keyword) =>
        lowerLine.StartsWith(keyword, StringComparison.Ordinal) &&
        (lowerLine.Length == keyword.Length || !char.IsLetterOrDigit(lowerLine[keyword.Length]));

    // ---- Display switches ---------------------------------------------------------------

    /// <summary>The display state set by keyword lines, applied to the inline expressions that follow.</summary>
    private sealed class SwitchState
    {
        private int? _decimals;
        private string? _angles;
        private bool? _substitute;
        private string? _mode; // "#val" or "#noc"

        public bool Hidden { get; private set; }

        public bool Apply(string line)
        {
            var lower = line.Trim().ToLowerInvariant();
            if (IsKeyword(lower, "#round"))
            {
                var n = lower.Length > 6 ? lower[6..].Trim() : "default";
                _decimals = n == "default" ? null : int.Parse(n, CultureInfo.InvariantCulture);
                return true;
            }
            switch (lower.Split(' ')[0])
            {
                case "#deg" or "#rad" or "#gra": _angles = lower.Split(' ')[0]; return true;
                case "#nosub": _substitute = false; return true;
                case "#varsub": _substitute = true; return true;
                case "#val": _mode = "#val"; return true;
                case "#noc": _mode = "#noc"; return true;
                case "#equ": _mode = null; return true;
                case "#hide" or "#pre": Hidden = true; return true;
                case "#show" or "#post": Hidden = false; return true;
                case "#wrap" or "#complex" or "#local" or "#global": return true; // defaults, or no effect here
                default: return false;
            }
        }

        public string Prefix(string expression)
        {
            if (Hidden)
                return "#hide " + expression;

            var sb = new StringBuilder();
            if (_mode is not null) sb.Append(_mode).Append(' ');
            if (_decimals is { } d) sb.Append("#round ").Append(d.ToString(CultureInfo.InvariantCulture)).Append(' ');
            if (_angles is not null) sb.Append(_angles).Append(' ');
            if (_substitute is { } s) sb.Append(s ? "#varsub " : "#nosub ");
            return sb.Append(expression).ToString();
        }

        /// <summary>Keyword lines that put a block in the same display state.</summary>
        public IEnumerable<string> KeywordLines()
        {
            if (_decimals is { } d) yield return "#round " + d.ToString(CultureInfo.InvariantCulture);
            if (_angles is not null) yield return _angles;
            if (_substitute is { } s) yield return s ? "#varsub" : "#nosub";
            if (_mode is not null) yield return _mode;
            if (Hidden) yield return "#hide";
        }
    }
}
