using System.Security.Cryptography;
using System.Text;

namespace CalcpadWord.Engine;

public enum ResultKind
{
    /// <summary>An inline expression rendered as an equation.</summary>
    Math,
    /// <summary>An inline <c>#hide</c> expression: calculated, nothing shown.</summary>
    Hidden,
    /// <summary>An inline expression or a block that could not run; the source stays editable.</summary>
    Error,
    /// <summary>A block's output (it may contain Calcpad error lines; see <see cref="EvaluationResult.ErrorCount"/>).</summary>
    Block,
}

/// <param name="Source">The source as written in the document (switches included).</param>
/// <param name="Omml">Office Math content for <see cref="ResultKind.Math"/> (no &lt;m:oMath&gt; wrapper).</param>
/// <param name="Text">Plain-text rendering, for logs and tests.</param>
/// <param name="Error">For errors: the message. For blocks: the first Calcpad error in the output, if any.</param>
public sealed record EvaluationResult(ResultKind Kind, string Source, string Omml, string Text, string? Error)
{
    /// <summary>For <see cref="ResultKind.Block"/>: the output to insert.</summary>
    public BlockOutput? Block { get; init; }

    /// <summary>For <see cref="ResultKind.Block"/>: how many Calcpad errors its output contains.</summary>
    public int ErrorCount { get; init; }

    /// <summary>
    /// For blocks: a hash of everything the output depends on (settings, every line before the
    /// block, the block itself). Equal hashes mean equal output, so it is the block's fingerprint.
    /// </summary>
    public string? InputHash { get; init; }

    /// <summary>
    /// A block whose inputs match the hash it was last drawn with: its output in the document is
    /// current, so it was only re-run hidden (for its variables) and <see cref="Block"/> is null.
    /// </summary>
    public bool Unchanged { get; init; }

    public static EvaluationResult Failure(string source, string message) =>
        new(ResultKind.Error, source, string.Empty, string.Empty, message);

    public static EvaluationResult HiddenResult(string source, string text) =>
        new(ResultKind.Hidden, source, string.Empty, text, null);

    /// <summary>Short hash of everything that affects the rendering. Unchanged results are not redrawn.</summary>
    public string Fingerprint
    {
        get
        {
            if (InputHash is not null)
                return InputHash;

            var sb = new StringBuilder()
                .Append(Kind).Append('\u001F').Append(Source).Append('\u001F')
                .Append(Omml).Append('\u001F').Append(Text).Append('\u001F').Append(Error);
            if (Block is not null)
            {
                sb.Append('\u001F').Append(Block.Html);
                foreach (var e in Block.Expressions)
                    sb.Append('\u001F').Append(e);
            }
            return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString())), 0, 8);
        }
    }
}
