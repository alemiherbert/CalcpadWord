using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Calcpad.Core;

namespace CalcpadWord.Engine;

/// <param name="Kind">Inline expression, block or input field.</param>
/// <param name="Source">The source: one expression (with switches), a block's lines joined with '\n', or an input field's value.</param>
/// <param name="KnownInputHash">
/// For a block showing output: the <see cref="EvaluationResult.InputHash"/> it was drawn with.
/// If its inputs still hash the same, the (possibly slow) block is not drawn again.
/// </param>
/// <param name="Field">For an input field: the variable it sets.</param>
public sealed record WorksheetItem(ControlKind Kind, string Source, string? KnownInputHash = null, InputField? Field = null)
{
    public static WorksheetItem Inline(string source) => new(ControlKind.Inline, source);
    public static WorksheetItem Block(string source, string? knownInputHash = null) => new(ControlKind.Block, source, knownInputHash);
    public static WorksheetItem Input(InputField field, string value) => new(ControlKind.Input, value, null, field);
}

/// <summary>
/// Evaluates a whole document as one Calcpad worksheet: every inline expression becomes a line,
/// every block contributes its lines, in document order. Calcpad's ExpressionParser runs once,
/// so blocks and inline expressions share variables, functions, units and macros.
/// <para>
/// A marker comment line is put in front of every item and Calcpad's HTML output is split at
/// those markers, which keeps the mapping exact even when macros expand lines. Per-expression
/// switches become keyword lines around the expression, and after each block the state a block
/// can change (#hide, #round, #md, …) is reset to the document settings.
/// </para>
/// </summary>
public static partial class Worksheet
{
    private const string MarkerWord = "CPWITEM";

    [GeneratedRegex(@"<p[^>]*>\s*" + MarkerWord + @"(\d+)" + MarkerWord + @"\s*</p>\s*")]
    private static partial Regex Marker();

    // Function definitions have no result, so #val shows the whole definition.
    [GeneratedRegex(@"^\s*[\p{L}_][\p{L}\p{N}_\u2032\u2033\u2034\u2057,.]*\s*\([^()]*\)\s*=(?!=)")]
    private static partial Regex FunctionDefinition();

    // Output that depends on more than the script text: files, and random numbers.
    [GeneratedRegex(@"^\s*#(include|read)\b|\brandom\s*\(", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
    private static partial Regex NotRepeatable();

    /// <summary>Part of every input hash: bump when block rendering changes, so existing blocks are redrawn.</summary>
    private const string RenderVersion = "calcpad-word-render-2"; // 2: scientific notation (× and exponent minus)

    static Worksheet()
    {
        // Error messages are parsed, so keep them in English whatever Word's UI language is.
        Messages.Culture = CultureInfo.InvariantCulture;
    }

    /// <param name="baseDirectory">Folder that relative <c>#include</c> paths are resolved against (the document's folder).</param>
    /// <param name="cancellationToken">Stops Calcpad (it checks between steps, also inside plots and solvers).</param>
    public static IReadOnlyList<EvaluationResult> Evaluate(
        IReadOnlyList<WorksheetItem> items, DocumentSettings settings, string? baseDirectory = null,
        CancellationToken cancellationToken = default)
    {
        var results = new EvaluationResult?[items.Count];
        var directives = new InlineDirective?[items.Count];
        var inputHashes = new string?[items.Count];
        var reset = StateLines(settings);
        var script = new StringBuilder();

        // Hash of everything emitted so far (markers excluded: they only number the items).
        using var inputs = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Hash(inputs, RenderVersion, typeof(ExpressionParser).Assembly.GetName().Version?.ToString(), settings.Serialize());
        var repeatable = true;

        AppendLines(script, reset);
        Hash(inputs, reset);

        for (var i = 0; i < items.Count; ++i)
        {
            var item = items[i];
            if (item.Kind == ControlKind.Inline)
            {
                InlineDirective directive;
                try
                {
                    directive = InlineDirective.Parse(SourceNormalizer.Normalize(item.Source));
                }
                catch (FormatException ex)
                {
                    results[i] = EvaluationResult.Failure(item.Source, ex.Message);
                    continue;
                }
                if (InlineProblem(directive.Expression) is { } problem)
                {
                    results[i] = EvaluationResult.Failure(item.Source, problem);
                    continue;
                }
                directives[i] = directive;
                AppendMarker(script, i);
                var (on, off) = SwitchLines(directive, settings);
                List<string> lines = [.. on, directive.Expression, .. off];
                AppendLines(script, lines);
                Hash(inputs, lines);
                repeatable &= !NotRepeatable().IsMatch(directive.Expression);
            }
            else if (item.Kind == ControlKind.Input)
            {
                // A form field: its value becomes a hidden assignment. The field itself shows the value.
                string assignment;
                try
                {
                    assignment = item.Field!.Assignment(item.Source);
                }
                catch (FormatException ex)
                {
                    results[i] = EvaluationResult.Failure(item.Source, ex.Message);
                    continue;
                }
                AppendMarker(script, i);
                List<string> lines = ["#hide", assignment, "#show"];
                AppendLines(script, lines);
                Hash(inputs, lines);
            }
            else
            {
                var lines = BlockSource.Lines(SourceNormalizer.NormalizeBlock(item.Source));
                if (BlockSource.Problem(lines) is { } problem)
                {
                    results[i] = EvaluationResult.Failure(item.Source, problem);
                    continue;
                }
                var resolved = lines.Select(line => ResolveInclude(line, baseDirectory)).ToList();
                Hash(inputs, resolved);
                repeatable &= !resolved.Any(line => NotRepeatable().IsMatch(line));
                var hash = inputHashes[i] = Convert.ToHexString(inputs.GetCurrentHash(), 0, 8);

                AppendMarker(script, i);
                if (repeatable && item.KnownInputHash == hash)
                {
                    // Its output in the document is current. Run it hidden, which skips plots and
                    // maps but still defines its variables for the rest of the document.
                    results[i] = new EvaluationResult(ResultKind.Block, item.Source, string.Empty, string.Empty, null)
                    {
                        InputHash = hash,
                        Unchanged = true,
                    };
                    script.Append("#hide\n");
                }
                AppendLines(script, resolved);
                AppendLines(script, reset); // starts with #show
                Hash(inputs, reset);
            }
        }

        var code = ExpandMacros(script.ToString(), out var macroErrors);
        var parser = new ExpressionParser
        {
            Settings = new Settings
            {
                Math = new MathSettings
                {
                    Decimals = settings.Decimals,
                    Substitute = true, // the document setting is applied with #nosub / #varsub
                    Degrees = (int)settings.Angles,
                    FormatEquations = true,
                },
            },
            ShowWarnings = true,
        };
        using (cancellationToken.Register(parser.Cancel))
            parser.Parse(code, calculate: true, getXml: true);
        cancellationToken.ThrowIfCancellationRequested();
        var segments = SplitAtMarkers(parser.HtmlResult);
        var expressions = parser.OpenXmlExpressions;

        for (var i = 0; i < items.Count; ++i)
        {
            if (results[i] is not null)
                continue;

            var source = items[i].Source;
            if (macroErrors.TryGetValue(i, out var macroError))
                results[i] = EvaluationResult.Failure(source, macroError);
            else if (!segments.TryGetValue(i, out var html))
                results[i] = EvaluationResult.Failure(source, "Not calculated: Calcpad stopped earlier in the document (too many errors?).");
            else if (items[i].Kind == ControlKind.Input)
                results[i] = CalcpadHtml.Errors(CalcpadHtml.Parse(html)) is [var error, ..]
                    ? EvaluationResult.Failure(source, CalcpadHtml.ErrorText(error))
                    : EvaluationResult.HiddenResult(source, items[i].Field!.Assignment(source));
            else if (items[i].Kind == ControlKind.Inline)
                results[i] = InlineResult(source, directives[i]!, html, expressions, settings);
            else
                results[i] = BlockResult(source, html, expressions, settings) with { InputHash = inputHashes[i] };
        }
        return results!;
    }

    private static void Hash(IncrementalHash hash, params IEnumerable<string?> lines)
    {
        foreach (var line in lines)
        {
            hash.AppendData(Encoding.UTF8.GetBytes(line ?? string.Empty));
            hash.AppendData("\n"u8);
        }
    }

    // ---- Script building ----------------------------------------------------------------

    /// <summary>Keyword lines that put Calcpad in the document's default state.</summary>
    private static string[] StateLines(DocumentSettings s) =>
    [
        "#show", "#equ", s.Substitute ? "#varsub" : "#nosub", "#wrap", "#round default",
        "#format default", "#md off", "#complex", "#" + DocumentSettings.AngleKey(s.Angles),
    ];

    private static (List<string> On, List<string> Off) SwitchLines(InlineDirective d, DocumentSettings s)
    {
        List<string> on = [], off = [];
        if (d.Decimals is { } decimals)
        {
            on.Add($"#round {decimals.ToString(CultureInfo.InvariantCulture)}");
            off.Add("#round default");
        }
        if (d.Substitute is { } substitute && substitute != s.Substitute)
        {
            on.Add(substitute ? "#varsub" : "#nosub");
            off.Add(s.Substitute ? "#varsub" : "#nosub");
        }
        if (d.Angles is { } angles && angles != s.Angles)
        {
            on.Add("#" + DocumentSettings.AngleKey(angles));
            off.Add("#" + DocumentSettings.AngleKey(s.Angles));
        }
        if (d.Mode == DisplayMode.NoCalculation)
        {
            on.Add("#noc");
            off.Add("#equ");
        }
        else if (d.Mode == DisplayMode.Hidden)
        {
            on.Add("#hide");
            off.Add("#show");
        }
        return (on, off);
    }

    private static string? InlineProblem(string expression)
    {
        if (expression.Length == 0)
            return "There is no expression to evaluate.";
        if (expression.Contains('\'') || expression.Contains('"'))
            return "Quotes start Calcpad comments. Write text outside the expression, or use a Calcpad block.";
        if (expression.StartsWith("$plot", StringComparison.OrdinalIgnoreCase) ||
            expression.StartsWith("$map", StringComparison.OrdinalIgnoreCase))
            return "Plots need a Calcpad block.";
        if (BlockSource.EndsWithContinuation(expression))
            return "The expression is incomplete.";
        return null;
    }

    private static void AppendMarker(StringBuilder script, int index) =>
        script.Append('\'').Append(MarkerWord).Append(index).Append(MarkerWord).Append('\n');

    private static void AppendLines(StringBuilder script, IEnumerable<string> lines)
    {
        foreach (var line in lines)
            script.Append(line).Append('\n');
    }

    /// <summary>Makes a relative <c>#include</c> path relative to the document's folder.</summary>
    private static string ResolveInclude(string line, string? baseDirectory)
    {
        var trimmed = line.TrimStart();
        if (baseDirectory is null || !trimmed.StartsWith("#include ", StringComparison.OrdinalIgnoreCase))
            return line;

        var rest = trimmed[9..];
        var end = rest.IndexOfAny(['\'', '"', '#']);
        var path = (end < 0 ? rest : rest[..end]).Trim();
        if (path.Length == 0 || Path.IsPathRooted(path) || path.StartsWith('%'))
            return line;

        return "#include " + Path.Combine(baseDirectory, path) + (end < 0 ? "" : " " + rest[end..]);
    }

    // ---- Macros -------------------------------------------------------------------------

    /// <summary>
    /// Expands #def / #include. MacroParser reports problems as lines of text in its output;
    /// those are removed and attributed to the item they belong to.
    /// </summary>
    private static string ExpandMacros(string script, out Dictionary<int, string> errors)
    {
        errors = [];
        var macroParser = new MacroParser { Include = IncludeFile };
        if (!macroParser.Parse(script, out var expanded, null, 0, false))
            return expanded;

        var errorStart = Messages.Error_in_0_on_line_1_2[..Messages.Error_in_0_on_line_1_2.IndexOf('{')];
        var output = new StringBuilder(expanded.Length);
        var item = -1;
        foreach (var line in expanded.Split('\n'))
        {
            var text = line.TrimEnd('\r');
            if (TryReadMarker(text, out var index))
                item = index;

            if (text.StartsWith(errorStart, StringComparison.Ordinal))
            {
                if (item >= 0 && !errors.ContainsKey(item))
                    errors[item] = CalcpadHtml.ErrorText(CalcpadHtml.Parse(text));
                continue;
            }
            output.Append(text).Append('\n');
        }
        return output.ToString();
    }

    private static bool TryReadMarker(string line, out int index)
    {
        index = -1;
        var t = line.Trim();
        if (!t.StartsWith("'" + MarkerWord, StringComparison.Ordinal) || !t.EndsWith(MarkerWord, StringComparison.Ordinal))
            return false;

        return int.TryParse(t.AsSpan(MarkerWord.Length + 1, t.Length - 2 * MarkerWord.Length - 1), out index);
    }

    /// <summary>Content of an included Calcpad file (input-form data after '\v' and #local sections are left out).</summary>
    private static string IncludeFile(string fileName, Queue<string> fields)
    {
        var text = File.ReadAllText(fileName);
        var formStart = text.IndexOf('\v');
        if (formStart >= 0)
            text = text[..formStart];

        var lines = new List<string>();
        var isLocal = false;
        foreach (var line in text.Split('\n'))
        {
            var t = line.Trim().ToLowerInvariant();
            if (t.StartsWith("#local", StringComparison.Ordinal)) isLocal = true;
            else if (t.StartsWith("#global", StringComparison.Ordinal)) isLocal = false;
            else if (!isLocal) lines.Add(line.TrimEnd('\r'));
        }
        return string.Join('\n', lines);
    }

    // ---- Reading the output -------------------------------------------------------------

    private static Dictionary<int, string> SplitAtMarkers(string html)
    {
        var segments = new Dictionary<int, string>();
        var matches = Marker().Matches(html);
        for (var m = 0; m < matches.Count; ++m)
        {
            var start = matches[m].Index + matches[m].Length;
            var end = m + 1 < matches.Count ? matches[m + 1].Index : html.Length;
            segments[int.Parse(matches[m].Groups[1].Value, CultureInfo.InvariantCulture)] = html[start..end];
        }
        return segments;
    }

    private static EvaluationResult InlineResult(
        string source, InlineDirective directive, string html, IReadOnlyList<string> expressions, DocumentSettings settings)
    {
        var root = CalcpadHtml.Parse(html);
        if (CalcpadHtml.Errors(root) is [var error, ..])
            return EvaluationResult.Failure(source, CalcpadHtml.ErrorText(error));

        var equation = CalcpadHtml.Equations(root).FirstOrDefault();
        if (equation is null)
        {
            return directive.Mode == DisplayMode.Hidden
                ? EvaluationResult.HiddenResult(source, directive.Expression)
                : EvaluationResult.Failure(source, "Calcpad produced no result for this expression.");
        }

        var index = CalcpadHtml.EquationIndex(equation);
        if (index < 0 || index >= expressions.Count)
            return EvaluationResult.Failure(source, "Calcpad produced no equation for this expression.");

        var omml = expressions[index];
        var text = CalcpadHtml.Text(equation);
        if (directive.Mode == DisplayMode.Value && !FunctionDefinition().IsMatch(directive.Expression))
        {
            omml = WordMarkup.ResultOnly(omml);
            text = WordMarkup.TextResultOnly(text);
        }
        return new EvaluationResult(ResultKind.Math, source, WordMarkup.ApplyStyle(omml, settings.Style, settings.MathFont), text, null);
    }

    private static EvaluationResult BlockResult(
        string source, string html, IReadOnlyList<string> expressions, DocumentSettings settings)
    {
        // Renumber equations from 0 so an unchanged block keeps an unchanged fingerprint.
        var local = new List<string>();
        var blockHtml = CalcpadHtml.EquationId().Replace(html, m =>
        {
            var k = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            local.Add(k < expressions.Count ? WordMarkup.ApplyStyle(expressions[k], settings.Style, settings.MathFont) : string.Empty);
            return $"id=\"eq-{local.Count - 1}\"";
        });
        blockHtml = CalcpadHtml.LineReference().Replace(blockHtml, string.Empty);

        var root = CalcpadHtml.Parse(blockHtml);
        var errors = CalcpadHtml.Errors(root);
        return new EvaluationResult(ResultKind.Block, source, string.Empty, CalcpadHtml.Text(root),
            errors.Count == 0 ? null : CalcpadHtml.ErrorText(errors[0]))
        {
            Block = new BlockOutput(blockHtml, local),
            ErrorCount = errors.Count,
        };
    }
}
