using System.Runtime.InteropServices;
using CalcpadWord.Engine;
using Word = Microsoft.Office.Interop.Word;

namespace CalcpadWord.AddIn;

internal sealed record CalcpadError(string Source, string Message, Word.ContentControl Control);

/// <summary>What is under the cursor, for the right-click menu.</summary>
internal enum CalcpadKind { None, Result, Source }

/// <summary>
/// Word-side logic. Every Calcpad expression or block in a document is a rich-text content
/// control whose Tag holds its source (see <see cref="ControlTag"/>). Recalculation collects
/// them in reading order (main text, text boxes, notes, then headers and footers) and runs
/// them as one Calcpad worksheet (<see cref="Worksheet"/>).
/// </summary>
internal sealed class DocumentController(Word.Application app)
{
    private const string SettingsVariable = "Calcpad.Settings";
    private const string ErrorTitlePrefix = "Error: ";
    private const int MaxTitle = 250;
    private const int MaxDelimiters = 10_000;

    // Order in which stories are evaluated. Comments and note separators are not scanned.
    private static readonly Word.WdStoryType[] StoryOrder =
    [
        Word.WdStoryType.wdMainTextStory,
        Word.WdStoryType.wdTextFrameStory,
        Word.WdStoryType.wdFootnotesStory,
        Word.WdStoryType.wdEndnotesStory,
        Word.WdStoryType.wdFirstPageHeaderStory,
        Word.WdStoryType.wdPrimaryHeaderStory,
        Word.WdStoryType.wdEvenPagesHeaderStory,
        Word.WdStoryType.wdFirstPageFooterStory,
        Word.WdStoryType.wdPrimaryFooterStory,
        Word.WdStoryType.wdEvenPagesFooterStory,
    ];

    private bool _busy;

    // ---- Commands -----------------------------------------------------------------------

    /// <summary>
    /// Turns the selection (or the current paragraph) into Calcpad, then recalculates the
    /// document. A selection that reads like a script (comments, headings, keywords, plots)
    /// becomes one block; otherwise each paragraph becomes one inline expression. Inside an
    /// existing expression or block it just recalculates.
    /// </summary>
    public string Evaluate()
    {
        var doc = ActiveDocument();
        using var undo = new UndoScope(app, "Calcpad Convert");
        var selection = app.Selection.Range;
        var note = string.Empty;
        var existing = FindCalcpadControl(selection);
        if (existing is null)
        {
            var paragraphs = WholeParagraphs(selection);
            if (paragraphs.ContentControls.Count == 0 && BlockSource.LooksLikeScript(ParagraphTexts(paragraphs)))
            {
                // A pasted Calcpad script: headings and text become Word paragraphs, calculations Calcpad.
                var source = SourceNormalizer.NormalizeBlock(paragraphs.Text);
                var conversion = ScriptConverter.Convert(source);
                if (conversion.Converted)
                {
                    InsertConverted(paragraphs, conversion);
                }
                else
                {
                    CreateBlock(paragraphs, source);
                    note = $" Kept the script as one Calcpad block because {conversion.KeptAsBlockReason}.";
                }
            }
            else
            {
                WrapAsExpressions(selection);
            }
        }
        return CalculateEdits(doc, keepActiveSource: false, existing?.ID) + note;
    }

    /// <summary>
    /// Replaces the block at the cursor with Word content: headings and text become Word
    /// paragraphs, expressions inline Calcpad expressions, and only plots, tables and control
    /// flow stay in (smaller) blocks.
    /// </summary>
    public string ConvertBlockToText()
    {
        var doc = ActiveDocument();
        var cc = FindCalcpadControl(app.Selection.Range);
        if (cc is null || ControlTag.Parse(cc.Tag) is not { Kind: ControlKind.Block } tag)
            throw new UserMessageException("Put the cursor in a Calcpad block first.");

        var source = tag.ShowsSource ? ReadEditedSource(cc, ControlKind.Block) : tag.Source;
        var conversion = ScriptConverter.Convert(source);
        if (!conversion.Converted)
            throw new UserMessageException($"This block stays a Calcpad block because {conversion.KeptAsBlockReason}.");

        using var undo = new UndoScope(app, "Calcpad Block to Text");
        _busy = true;
        try
        {
            // A range over the block's paragraphs follows them once the control is unwrapped.
            var content = cc.Range;
            var paragraphs = Slice(content, content.Paragraphs.First.Range.Start, content.Paragraphs.Last.Range.End);
            cc.LockContents = false;
            cc.Delete(false);
            InsertConverted(paragraphs, conversion);
        }
        finally
        {
            _busy = false;
        }
        return CalculateEdits(doc, keepActiveSource: false);
    }

    public bool IsBlockAt(Word.Range range) =>
        !_busy && FindCalcpadControl(range) is { } cc && ControlTag.Parse(cc.Tag)?.Kind == ControlKind.Block;

    /// <summary>Keyboard Evaluate: on a result it shows the source; anywhere else it evaluates.</summary>
    public string EvaluateOrEditSource() =>
        IsCalcpadResult(app.Selection.Range) ? ToggleEditSource() : Evaluate();

    /// <summary>
    /// Makes the selected paragraphs (or the current one) a Calcpad block. An empty paragraph
    /// becomes an empty block, ready to type into.
    /// </summary>
    public string MakeBlock()
    {
        var doc = ActiveDocument();
        var selection = app.Selection.Range;
        if (FindCalcpadControl(selection) is not null)
            throw new UserMessageException("The cursor is already in a Calcpad expression or block.");

        var paragraphs = WholeParagraphs(selection);
        if (paragraphs.ContentControls.Count > 0)
            throw new UserMessageException("The selection already contains content controls. Select ordinary paragraphs to make a block.");

        using var undo = new UndoScope(app, "Calcpad Block");
        var source = SourceNormalizer.NormalizeBlock(paragraphs.Text);
        var cc = CreateBlock(paragraphs, source);
        if (source.Length > 0)
            return CalculateEdits(doc, keepActiveSource: false);

        _busy = true;
        try
        {
            ShowSource(cc, string.Empty, TextLook.Source, "Calcpad block: type the script, then click outside it to calculate",
                ControlState.Editing, string.Empty);
            PlaceCursorAtEnd(cc);
        }
        finally
        {
            _busy = false;
        }
        return "Type the Calcpad script in the block, one statement per line. Click outside it to calculate.";
    }

    public string Recalculate() => Recalculate(ActiveDocument(), keepActiveSource: false);

    /// <param name="keepActiveSource">
    /// Leave the source the cursor is in as it is (used when another control was just left,
    /// so the user is not thrown out of the source they moved into).
    /// </param>
    public string Recalculate(Word.Document doc, bool keepActiveSource)
    {
        using var undo = new UndoScope(app, "Calcpad Calculate");
        return RecalculateCore(doc, keepActiveSource);
    }

    /// <summary>Switches the expression or block at the cursor between its result and its editable source.</summary>
    public string ToggleEditSource()
    {
        var doc = ActiveDocument();
        var cc = FindCalcpadControl(app.Selection.Range)
            ?? throw new UserMessageException("Put the cursor in a Calcpad result first.");
        var tag = ControlTag.Parse(cc.Tag)!;

        using var undo = new UndoScope(app, "Calcpad Edit Source");
        if (tag.ShowsSource)
            return CalculateEdits(doc, keepActiveSource: false, cc.ID);

        _busy = true;
        try
        {
            ShowSource(cc, tag.Source, TextLook.Source,
                tag.Kind == ControlKind.Block
                    ? "Calcpad block: click outside to calculate"
                    : "Calcpad source: Enter or click outside to calculate; Shift+Enter adds an expression",
                ControlState.Editing, string.Empty);
            PlaceCursorAtEnd(cc);
        }
        finally
        {
            _busy = false;
        }
        return tag.Kind == ControlKind.Block
            ? "Editing the Calcpad block. Click outside it to calculate."
            : "Editing the Calcpad source. Enter or click outside to calculate; Shift+Enter adds another expression.";
    }

    /// <summary>
    /// True when leaving <paramref name="cc"/> should recalculate: it shows source that was edited,
    /// or it is an input field whose value differs from the last calculation's.
    /// </summary>
    public bool NeedsEvaluationAfterExit(Word.ContentControl cc)
    {
        if (_busy)
            return false;
        if (InputField.IsInput(cc.Tag))
            return !_inputValues.TryGetValue(cc.ID, out var last) || last != ReadInput(cc);
        if (ControlTag.Parse(cc.Tag) is not { } tag)
            return false;

        return tag.State switch
        {
            ControlState.Editing => true,
            ControlState.Error => ReadEditedSource(cc, tag.Kind) != tag.Source, // an untouched error would fail again
            _ => false,
        };
    }

    /// <summary>
    /// True where Enter means "done here, carry on writing": in an inline expression (its source
    /// or its result) and in a block's output. In a block's script Enter adds a line, as usual.
    /// </summary>
    public bool EnterLeavesControlAt(Word.Range range)
    {
        if (_busy || FindCalcpadControl(range) is not { } cc || ControlTag.Parse(cc.Tag) is not { } tag)
            return false;

        return tag.Kind == ControlKind.Inline || !tag.ShowsSource;
    }

    /// <summary>
    /// Enter in an expression: leave it (an edited source is then evaluated by the exit event)
    /// and, if nothing else follows on its line, start a new paragraph below. Inside a sentence
    /// the cursor just moves past the expression. After a block's output, a new paragraph
    /// follows the block. Returns false if the cursor is not where Enter means that.
    /// </summary>
    public bool LeaveWithEnter()
    {
        if (TryActiveDocument() is null || !EnterLeavesControlAt(app.Selection.Range))
            return false;

        var cc = FindCalcpadControl(app.Selection.Range)!;
        var tag = ControlTag.Parse(cc.Tag)!;
        using var undo = new UndoScope(app, "Calcpad Enter");
        if (tag.Kind == ControlKind.Block)
        {
            var next = cc.Range.Paragraphs.Last.Next();
            if (next is null)
                return false;

            var start = next.Range.Start;
            Slice(next.Range, start, start).InsertParagraphBefore();
            Slice(next.Range, start, start).Select();
            return true;
        }

        var after = cc.Range.End + 1; // just past the control's closing boundary
        var paragraph = cc.Range.Paragraphs.Last.Range;
        var rest = Slice(paragraph, after, paragraph.End).Text ?? string.Empty;
        if (rest.Trim(' ', '\t', '\u00A0', '\r', '\a').Length == 0)
        {
            Slice(paragraph, after, after).InsertParagraphAfter();
            Slice(paragraph, after + 1, after + 1).Select();
        }
        else
        {
            Slice(paragraph, after, after).Select();
        }
        return true;
    }

    public bool IsCalcpadResult(Word.Range range) => KindAt(range) == CalcpadKind.Result;

    public CalcpadKind KindAt(Word.Range range)
    {
        if (_busy || FindCalcpadControl(range) is not { } cc || ControlTag.Parse(cc.Tag) is not { } tag)
            return CalcpadKind.None;

        return tag.ShowsSource ? CalcpadKind.Source : CalcpadKind.Result;
    }

    public DocumentSettings GetSettings() =>
        TryActiveDocument() is { } doc ? LoadSettings(doc) : new DocumentSettings();

    public string ChangeSettings(Func<DocumentSettings, DocumentSettings> change)
    {
        var doc = ActiveDocument();
        var before = LoadSettings(doc);
        var after = change(before);
        SaveSettings(doc, after);
        if (!after.Automatic)
        {
            if (before.Automatic)
                return LastStatus = "Calcpad: automatic calculation is off. Calculate (Ctrl+Alt+F9) updates the results.";
            return after.ResultKey() == before.ResultKey() ? LastStatus : MarkOutOfDate(doc, left: null);
        }
        using var undo = new UndoScope(app, "Calcpad Settings");
        return RecalculateCore(doc, keepActiveSource: false);
    }

    /// <summary>
    /// Leaving a changed input field or edited source: recalculates. With automatic calculation
    /// off (as in Excel's manual mode), a changed input field only marks the document's results
    /// out of date, so typing inputs stays instant, and edited source is calculated and drawn on
    /// its own, so it never stays open; what follows it waits for Calculate.
    /// </summary>
    public string AfterExit(Word.Document doc, Word.ContentControl left)
    {
        if (InputField.IsInput(left.Tag) && !LoadSettings(doc).Automatic)
            return MarkOutOfDate(doc, left);

        string? leftId = null;
        try { leftId = left.ID; }
        catch (COMException) { } // deleted meanwhile
        using var undo = new UndoScope(app, "Calcpad Calculate");
        return CalculateEdits(doc, keepActiveSource: true, leftId);
    }

    private const string OutOfDateVariable = "Calcpad.OutOfDate";

    /// <summary>
    /// Automatic calculation off, an input changed: the document's results are marked out of date (a document
    /// variable, kept when it is saved; the Calculate button says so) until the next
    /// Calculate draws them. Nothing in the text changes, so leaving a field is instant.
    /// Greying each result instead took about 3 s on the v2 beam.
    /// </summary>
    private string MarkOutOfDate(Word.Document doc, Word.ContentControl? left)
    {
        if (left is not null && InputField.IsInput(left.Tag))
            _inputValues[left.ID] = ReadInput(left); // leaving it again unchanged is no change
        SetOutOfDate(doc, true);
        return LastStatus = "Calcpad: results out of date. Calculate (Ctrl+Alt+F9) updates them.";
    }

    /// <summary>The active document's results are out of date: changed with automatic calculation off, not recalculated since.</summary>
    public bool IsOutOfDate() => TryActiveDocument() is { } doc && IsOutOfDate(doc);

    private static bool IsOutOfDate(Word.Document doc)
    {
        try
        {
            return doc.Variables[OutOfDateVariable].Value == "1";
        }
        catch (COMException)
        {
            return false; // never marked
        }
    }

    private static void SetOutOfDate(Word.Document doc, bool outOfDate)
    {
        if (IsOutOfDate(doc) == outOfDate)
            return;
        try
        {
            if (outOfDate)
                doc.Variables.Add(OutOfDateVariable, "1");
            else
                doc.Variables[OutOfDateVariable].Delete();
        }
        catch (COMException ex)
        {
            Log.Error("Marking the results out of date failed", ex); // only a hint
        }
    }

    // ---- Errors -------------------------------------------------------------------------

    public IReadOnlyList<CalcpadError> Errors() =>
        TryActiveDocument() is { } doc ? Errors(doc) : [];

    /// <summary>Input fields with values that can't be used, expressions and blocks that could not run, and blocks whose output contains errors.</summary>
    public static IReadOnlyList<CalcpadError> Errors(Word.Document doc)
    {
        var errors = new List<CalcpadError>();
        foreach (var cc in WorksheetControls(doc))
        {
            if (InputField.Parse(cc.Tag) is { } field)
            {
                try
                {
                    field.Assignment(ReadInput(cc));
                }
                catch (FormatException ex)
                {
                    errors.Add(new CalcpadError("Input " + field.Name, ex.Message, cc));
                }
                continue;
            }

            var tag = ControlTag.Parse(cc.Tag)!;
            var title = TitleOf(cc);
            var hasError = tag.State == ControlState.Error ||
                           tag.Kind == ControlKind.Block && title.StartsWith(ErrorTitlePrefix, StringComparison.Ordinal);
            if (!hasError)
                continue;

            var message = title.StartsWith(ErrorTitlePrefix, StringComparison.Ordinal) ? title[ErrorTitlePrefix.Length..] : title;
            var label = tag.Kind == ControlKind.Block ? "Block: " + FirstLine(tag.Source) : tag.Source;
            errors.Add(new CalcpadError(label, message, cc));
        }
        return errors;
    }

    /// <summary>Moves the cursor to the next error after the cursor, wrapping around.</summary>
    public string NextError()
    {
        var doc = ActiveDocument();
        var errors = Errors(doc);
        if (errors.Count == 0)
            return "Calcpad: no errors.";

        var selection = app.Selection.Range;
        var index = 0;
        if (FindCalcpadControl(selection) is { } current &&
            errors.Select(e => e.Control.ID).ToList().IndexOf(current.ID) is var i and >= 0)
        {
            index = (i + 1) % errors.Count;
        }
        else
        {
            var after = errors.Select((e, n) => (e, n)).FirstOrDefault(x =>
                x.e.Control.Range.StoryType == selection.StoryType && x.e.Control.Range.Start > selection.Start);
            index = after.e is null ? 0 : after.n;
        }
        return GoToError(errors, index);
    }

    public string GoToError(int index)
    {
        var errors = Errors(ActiveDocument());
        if (errors.Count == 0)
            return "Calcpad: no errors.";

        return GoToError(errors, Math.Clamp(index, 0, errors.Count - 1));
    }

    private static string GoToError(IReadOnlyList<CalcpadError> errors, int index)
    {
        var error = errors[index];
        var at = error.Control.Range;
        var isBlock = ControlTag.Parse(error.Control.Tag)?.Kind == ControlKind.Block;
        at.Collapse(isBlock ? Word.WdCollapseDirection.wdCollapseStart : Word.WdCollapseDirection.wdCollapseEnd);
        at.Select();
        return $"Calcpad error {index + 1} of {errors.Count}: {error.Message}";
    }

    // ---- Recalculation ------------------------------------------------------------------

    // ---- Background calculation ----------------------------------------------------------
    //
    // Recalculation has three steps. Reading the document and drawing results use Word, so they
    // run on Word's thread; running Calcpad in between does not, so it runs on a background
    // thread. If Calcpad is done within ForegroundMilliseconds the results are drawn straight
    // away (the usual case: one Undo step, no flicker). Otherwise Word stays usable and the
    // results are drawn when ready, provided the Calcpad items haven't changed meanwhile; if they
    // have, it recalculates. A newer recalculation cancels the one still running.
    //
    // 1.5 s, not less: greying the results first costs about 25 ms a result in Word, so for a
    // calculation of a second or so, waiting and drawing straight away is quicker than greying
    // and then drawing (the v2 beam: 0.7 s in Calcpad, 3 s to grey 110 results).

    private const int ForegroundMilliseconds = 1500;

    /// <summary>
    /// Results are greyed only if Calcpad is still running this long after the calculation went
    /// to the background. One that finishes sooner (the soil template, 1.7 s in all) is drawn
    /// without greying: grey that came and went within half a second flashed the page.
    /// </summary>
    private const int GreyAfterMilliseconds = 1000;

    /// <summary>The background calculation's results, until they are greyed, drawn or cancelled. Word's thread.</summary>
    private Snapshot? _toGrey;

    /// <summary>What a calculation started from: the controls and the sources they had.</summary>
    /// <param name="Controls">The Calcpad expressions and blocks (with <paramref name="Ids"/>, <paramref name="Tags"/>, <paramref name="Parsed"/>).</param>
    /// <param name="Items">What the worksheet runs, in document order: the controls and the input fields.</param>
    /// <param name="ItemOf">For each of <paramref name="Controls"/>: its index in <paramref name="Items"/>.</param>
    /// <param name="Inputs">The input fields, with the value calculated with.</param>
    /// <param name="Only">
    /// Automatic calculation off, after an edit: the IDs of the only controls to draw (those just
    /// written or edited). The worksheet then runs only as far as the last of them (<paramref name="Limit"/>
    /// items), and everything else waits for Calculate, as in Excel's manual mode.
    /// </param>
    /// <param name="LeftId">The control just left, if any (with <paramref name="Only"/>).</param>
    private sealed record Snapshot(
        Word.Document Doc, List<Word.ContentControl> Controls, List<string> Ids, List<string> Tags,
        List<ControlTag> Parsed, List<WorksheetItem> Items, List<int> ItemOf, List<InputSnapshot> Inputs,
        DocumentSettings Settings, string? Folder, double? TextWidth,
        HashSet<string>? Only = null, int Limit = int.MaxValue, string? LeftId = null);

    /// <param name="Changed">The value differs from the last calculation's (or this session hasn't calculated it yet).</param>
    private sealed record InputSnapshot(Word.ContentControl Control, string Id, InputField Field, string Value, int Item, bool Changed);

    /// <summary>The value each input field (by control ID) had in the last calculation: leaving a changed one recalculates.</summary>
    private readonly Dictionary<string, string> _inputValues = new(StringComparer.Ordinal);

    // Written on Word's thread only; IsWorking and LastStatus may be read from any thread (an
    // automation client polling them must not keep interrupting Word while it draws).
    private volatile CancellationTokenSource? _calculation;
    private volatile Task _lastCalculation = Task.CompletedTask; // Calcpad runs one calculation at a time
    private volatile string _lastStatus = string.Empty;

    /// <summary>A calculation (or font download) is running in the background, or its results are still to be drawn. Any thread.</summary>
    public bool IsWorking => _calculation is not null || _fontInstall is not null || !_lastCalculation.IsCompleted;

    /// <summary>
    /// Runs Calcpad once on a background thread at start-up. Loading and compiling it takes
    /// longer than the foreground budget, and the first real calculation would otherwise
    /// always be drawn late.
    /// </summary>
    public void WarmUp()
    {
        _lastCalculation = Task.Run(() =>
        {
            try
            {
                var timer = System.Diagnostics.Stopwatch.StartNew();
                var results = Worksheet.Evaluate(
                    [WorksheetItem.Inline("a = 2*3"), WorksheetItem.Block("'Warm-up\nb = sqrt(a)*1m\nf(x) = x^2\nc = f(b)")],
                    new DocumentSettings());
                foreach (var result in results)
                    _ = result.Block?.WordPackage;
                Log.Info($"Calcpad warmed up in {timer.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Log.Error("Warming up Calcpad failed", ex);
            }
        });
    }

    /// <summary>At shutdown: stops a calculation still running.</summary>
    public void CancelBackgroundWork()
    {
        _calculation?.Cancel();
        _calculation = null;
        _fontInstallCancel?.Cancel();
    }

    /// <summary>A background calculation or font download can be cancelled. Any thread (the keyboard hook asks).</summary>
    public bool CanCancel => _calculation is not null || _fontInstallCancel is { IsCancellationRequested: false };

    /// <summary>
    /// The Cancel button (or Esc): stops the background calculation, or the font download.
    /// Results already in the document stay as they were.
    /// </summary>
    public string Cancel()
    {
        var what = new List<string>();
        if (_calculation is { } calculation)
        {
            calculation.Cancel();
            _calculation = null;
            what.Add("the calculation");
        }
        if (_fontInstallCancel is { IsCancellationRequested: false } fontInstall)
        {
            fontInstall.Cancel();
            what.Add("the font download");
        }
        if (what.Count == 0)
            return "Calcpad: nothing to cancel.";

        GreyNow(); // cancelled before they were greyed: they may be out of date all the same
        Log.Info($"Cancelled {string.Join(" and ", what)}");
        var hint = what.Contains("the calculation")
            ? " Results that may be out of date stay greyed: Calculate (Ctrl+Alt+F9) updates them."
            : string.Empty;
        return LastStatus = $"Calcpad: cancelled {string.Join(" and ", what)}.{hint}";
    }

    /// <summary>The status of the last finished recalculation. Any thread.</summary>
    public string LastStatus { get => _lastStatus; private set => _lastStatus = value; }

    /// <summary>Raised on Word's thread when results calculated in the background have been drawn (or failed).</summary>
    public event Action<string>? BackgroundFinished;

    /// <summary>
    /// After writing or editing Calcpad: with automatic calculation on, recalculates everything;
    /// off, calculates and draws only what was just written or edited (<see cref="Snapshot.Only"/>).
    /// </summary>
    /// <param name="leftId">The control just left (an error that was edited is drawn too).</param>
    private string CalculateEdits(Word.Document doc, bool keepActiveSource, string? leftId = null) =>
        // A full calculation still running (Calculate) is not cut short: it is started again in full.
        LoadSettings(doc).Automatic || (_calculation is not null && !_calculatingOnlyEdits)
            ? RecalculateCore(doc, keepActiveSource)
            : RecalculateCore(doc, keepActiveSource, onlyEdited: true, leftId);

    /// <summary>The calculation running (<see cref="_calculation"/>) draws only what was edited.</summary>
    private bool _calculatingOnlyEdits;

    private string RecalculateCore(Word.Document doc, bool keepActiveSource, bool onlyEdited = false, string? leftId = null)
    {
        if (_busy)
            return "Calcpad is busy.";

        Snapshot snapshot;
        _busy = true;
        try
        {
            using var fullSpeed = new FullSpeedScope();
            using var screen = new ScreenUpdatingScope(app);
            var activeId = keepActiveSource ? FindCalcpadControl(app.Selection.Range)?.ID : null;
            SplitMultiLineExpressions(doc, activeId);
            foreach (var story in Stories(doc))
                ConvertDelimiters(story);
            snapshot = TakeSnapshot(doc);
            if (onlyEdited)
                snapshot = LimitToEdits(snapshot, activeId, leftId);
        }
        finally
        {
            _busy = false;
        }
        if (snapshot.Only is { Count: 0 })
            return LastStatus = IsOutOfDate(doc) ? "Calcpad: results out of date. Calculate (Ctrl+Alt+F9) updates them." : LastStatus;

        _calculation?.Cancel(); // its results would be out of date
        _toGrey = null;
        var cancel = new CancellationTokenSource();
        _calculation = cancel;
        _calculatingOnlyEdits = snapshot.Only is not null;
        var token = cancel.Token;
        var calculation = _lastCalculation.ContinueWith(
            _ => Calculate(snapshot, token), CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
        _lastCalculation = calculation;

        if (Task.WaitAny([calculation], ForegroundMilliseconds) == 0)
            return Finish(snapshot, calculation, cancel, keepActiveSource, inBackground: false) ?? LastStatus;

        _toGrey = snapshot;
        UiThread.WhenDone(Task.WhenAny(calculation, Task.Delay(GreyAfterMilliseconds)), () =>
        {
            if (ReferenceEquals(_toGrey, snapshot) && !calculation.IsCompleted)
                GreyNow();
        });
        app.StatusBar = "Calcpad: calculating… you can keep working.";
        Log.Info($"Calculating '{doc.Name}' in the background ({snapshot.Items.Count} items)");
        UiThread.WhenDone(calculation, () =>
        {
            var status = Finish(snapshot, calculation, cancel, keepActiveSource: true, inBackground: true);
            if (status is not null)
                BackgroundFinished?.Invoke(status);
        });
        return "Calcpad: calculating in the background…";
    }

    /// <summary>Greys the results waiting for it (<see cref="_toGrey"/>), as one Undo step: the user may have typed meanwhile.</summary>
    private void GreyNow()
    {
        if (_toGrey is not { } snapshot)
            return;
        _toGrey = null;
        if (!IsOpen(snapshot.Doc))
            return;
        using var undo = new UndoScope(app, "Calcpad: out of date");
        GreyWhileCalculating(snapshot);
    }

    /// <summary>The results may change while Calcpad runs: grey them until they are drawn again (or for good, if cancelled).</summary>
    private void GreyWhileCalculating(Snapshot snapshot)
    {
        var stale = StaleCandidates(snapshot);
        if (stale.Count == 0)
            return; // nothing drawn yet that could change: leave Word alone

        _busy = true;
        try
        {
            // No ScreenUpdating switch here: a few formatting changes don't flicker, and an extra
            // off/on just before background results are drawn coincided with Word hanging.
            using var fullSpeed = new FullSpeedScope();
            var marked = 0;
            foreach (var cc in stale)
            {
                try
                {
                    if (IsStale(cc))
                        continue;
                    SetStale(cc, stale: true);
                    ++marked;
                }
                catch (COMException) when (!IsAlive(cc))
                {
                    // Deleted since the calculation started (greying waits a moment): nothing to grey.
                }
            }
            if (marked > 0)
                Log.Info($"Greyed {marked} result{(marked == 1 ? "" : "s")} until the calculation finishes");
        }
        catch (COMException ex)
        {
            Log.Error("Greying out-of-date results failed", ex); // only a hint: the calculation goes on
        }
        finally
        {
            _busy = false;
        }
    }

    private Snapshot TakeSnapshot(Word.Document doc)
    {
        List<Word.ContentControl> controls = [];
        List<string> tags = [];
        List<ControlTag> parsed = [];
        List<WorksheetItem> items = [];
        List<int> itemOf = [];
        List<InputSnapshot> inputs = [];
        foreach (var cc in WorksheetControls(doc))
        {
            var tag = cc.Tag;
            if (InputField.Parse(tag) is { } field)
            {
                var value = ReadInput(cc);
                var id = cc.ID;
                var changed = !_inputValues.TryGetValue(id, out var last) || last != value;
                _inputValues[id] = value;
                inputs.Add(new InputSnapshot(cc, id, field, value, items.Count, changed));
                items.Add(WorksheetItem.Input(field, value));
                continue;
            }

            var p = ControlTag.Parse(tag)!;
            controls.Add(cc);
            tags.Add(tag);
            parsed.Add(p);
            itemOf.Add(items.Count);
            items.Add(new WorksheetItem(
                p.Kind,
                p.ShowsSource ? ReadEditedSource(cc, p.Kind) : p.Source,
                // A block showing output carries the input hash it was drawn with: unchanged blocks aren't redrawn.
                p.Kind == ControlKind.Block && !p.ShowsSource ? p.Fingerprint : null));
        }
        return new Snapshot(doc, controls, [.. controls.Select(cc => cc.ID)], tags, parsed, items, itemOf, inputs,
            LoadSettings(doc), DocumentFolder(doc), TextWidthOf(doc));
    }

    /// <summary>
    /// Automatic calculation off: only the sources being edited, the new expressions and the
    /// control just left are drawn (not the source the cursor is in). The worksheet runs top to
    /// bottom, so it needs to run only as far as the last of them.
    /// </summary>
    private static Snapshot LimitToEdits(Snapshot s, string? activeId, string? leftId)
    {
        var only = new HashSet<string>(StringComparer.Ordinal);
        var limit = 0;
        for (var i = 0; i < s.Controls.Count; ++i)
        {
            var (id, tag) = (s.Ids[i], s.Parsed[i]);
            if (id == activeId || !(tag.State == ControlState.Editing || tag.Fingerprint.Length == 0 || id == leftId))
                continue;
            only.Add(id);
            limit = Math.Max(limit, s.ItemOf[i] + 1);
        }
        return s with { Only = only, Limit = limit, LeftId = leftId };
    }

    /// <summary>
    /// The text width in points, which plots are made to fill: the last section's page less its
    /// margins and gutter, or its column width. Null if Word can't say: plots stay as drawn.
    /// </summary>
    private static double? TextWidthOf(Word.Document doc)
    {
        try
        {
            var setup = doc.Sections.Last.PageSetup;
            var columns = setup.TextColumns;
            return columns.Count > 1 ? columns.Width : setup.PageWidth - setup.LeftMargin - setup.RightMargin - setup.Gutter;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>Background thread: Calcpad, then the Word content of blocks to draw. No Word calls here.</summary>
    private static IReadOnlyList<EvaluationResult> Calculate(Snapshot snapshot, CancellationToken token)
    {
        using var fullSpeed = new FullSpeedScope();
        var timer = System.Diagnostics.Stopwatch.StartNew();
        var items = snapshot.Limit < snapshot.Items.Count ? snapshot.Items.GetRange(0, snapshot.Limit) : snapshot.Items;
        var results = Worksheet.Evaluate(items, snapshot.Settings, snapshot.Folder, token);
        var calculated = timer.Elapsed;
        foreach (var result in results)
        {
            token.ThrowIfCancellationRequested();
            _ = result.Block?.WordPackageFor(snapshot.TextWidth);
        }
        if (timer.ElapsedMilliseconds > ForegroundMilliseconds)
            Log.Info($"Calculated in {calculated.TotalSeconds:F1} s, Word content ready in {timer.Elapsed.TotalSeconds:F1} s");
        return results;
    }

    /// <summary>Word's thread: draws the results. Null when a newer calculation replaced this one.</summary>
    private string? Finish(Snapshot s, Task<IReadOnlyList<EvaluationResult>> calculation, CancellationTokenSource cancel,
        bool keepActiveSource, bool inBackground)
    {
        if (cancel.IsCancellationRequested)
        {
            if (inBackground)
                Log.Info("A background calculation was cancelled, or replaced by a newer one");
            return null;
        }
        if (ReferenceEquals(_toGrey, s))
            _toGrey = null; // done before it was greyed
        // This calculation counts as running (IsWorking) until its results are drawn and LastStatus is set.
        var drawLater = false;
        try
        {
            if (!IsOpen(s.Doc))
            {
                Log.Info("The document was closed before its calculation finished");
                return null;
            }
            if (calculation.IsFaulted)
            {
                Log.Error("Calculation failed", calculation.Exception);
                return LastStatus = $"Calcpad: the calculation failed: {calculation.Exception?.GetBaseException().Message}";
            }
            if (_busy)
            {
                // Another command is running: draw right after it.
                Log.Info("Background results wait for a running command");
                drawLater = true;
                UiThread.Post(() =>
                {
                    var status = Finish(s, calculation, cancel, keepActiveSource, inBackground);
                    if (status is not null)
                        BackgroundFinished?.Invoke(status);
                });
                return null;
            }
            return DrawResults(s, calculation.Result, keepActiveSource, inBackground);
        }
        finally
        {
            if (!drawLater && ReferenceEquals(_calculation, cancel))
                _calculation = null;
        }
    }

    /// <summary>Word's thread. Null when the document changed meanwhile and a new background calculation started.</summary>
    private string? DrawResults(Snapshot s, IReadOnlyList<EvaluationResult> results, bool keepActiveSource, bool inBackground)
    {
        using var fullSpeed = new FullSpeedScope(); // Word may be in the background while results arrive
        using var undo = inBackground ? new UndoScope(app, "Calcpad results") : null;
        _busy = true;
        try
        {
            using var screen = new ScreenUpdatingScope(app);
            var controls = s.Controls;
            if (inBackground)
            {
                // The user kept working: are these still the items that were calculated?
                controls = CalcpadControls(s.Doc);
                if (!StillTheSame(s, controls))
                {
                    _busy = false;
                    Log.Info($"'{s.Doc.Name}' changed during the calculation: calculating again");
                    var restarted = RecalculateCore(s.Doc, keepActiveSource: true, onlyEdited: s.Only is not null, s.LeftId);
                    return _calculation is null ? restarted : null; // still running: it reports when done
                }
            }
            var activeId = keepActiveSource ? FindCalcpadControl(app.Selection.Range)?.ID : null;
            return LastStatus = Draw(s, controls, results, activeId);
        }
        finally
        {
            _busy = false;
        }
    }

    private static bool StillTheSame(Snapshot s, List<Word.ContentControl> now)
    {
        if (now.Count != s.Controls.Count)
            return false;
        for (var i = 0; i < now.Count; ++i)
        {
            if (now[i].ID != s.Ids[i] || now[i].Tag != s.Tags[i])
                return false;
            if (s.Parsed[i].ShowsSource && ReadEditedSource(now[i], s.Parsed[i].Kind) != s.Items[s.ItemOf[i]].Source)
                return false;
        }

        // And the input fields hold the values calculated with.
        var inputs = InputControls(s.Doc);
        if (inputs.Count != s.Inputs.Count)
            return false;
        for (var k = 0; k < inputs.Count; ++k)
        {
            if (inputs[k].ID != s.Inputs[k].Id || ReadInput(inputs[k]) != s.Inputs[k].Value)
                return false;
        }
        return true;
    }

    /// <summary>Word's thread: redraws what changed. <paramref name="activeId"/> is a source the user is typing in: left alone.</summary>
    private static string Draw(Snapshot s, List<Word.ContentControl> controls, IReadOnlyList<EvaluationResult> results, string? activeId)
    {
        int errors = 0, redrawn = 0;
        string? firstError = null;
        var timer = System.Diagnostics.Stopwatch.StartNew();
        _insertTicks = _deleteTicks = 0;
        var inline = new List<InlineResult>();

        // Input fields come first in the count: a wrong value explains the errors after it.
        // (Drawing only what was edited, they wait for Calculate like everything else.)
        foreach (var input in s.Only is null ? s.Inputs : [])
        {
            var result = results[input.Item];
            var failed = result.Kind == ResultKind.Error;
            if (failed)
            {
                ++errors;
                firstError ??= result.Error;
            }
            MarkInput(input.Control, failed);
        }
        var inputsDone = timer.Elapsed;
        long blockTicks = 0;

        for (var i = 0; i < controls.Count; ++i)
        {
            if (s.Only is not null && !s.Only.Contains(s.Ids[i]))
                continue;
            var (cc, tag, result) = (controls[i], s.Parsed[i], results[s.ItemOf[i]]);
            // A block that wasn't redrawn keeps the error summary in its title.
            var title = TitleOf(cc);
            var keptError = result.Unchanged && title.StartsWith(ErrorTitlePrefix, StringComparison.Ordinal)
                ? title[ErrorTitlePrefix.Length..]
                : null;
            if (result.Kind == ResultKind.Error || result.ErrorCount > 0 || keptError is not null)
            {
                ++errors;
                firstError ??= $"{(tag.Kind == ControlKind.Block ? "block" : FirstLine(result.Source))}: {keptError ?? result.Error}";
            }

            if (tag.ShowsSource && activeId is not null && cc.ID == activeId)
                continue; // the user is typing in it; it is evaluated when they leave

            var fingerprint = result.Fingerprint;
            var state = result.Kind == ResultKind.Error ? ControlState.Error : ControlState.Result;
            if (tag.State == state && tag.Fingerprint == fingerprint)
            {
                // Unchanged since it was last drawn: only the out-of-date mark goes.
                if (IsStale(cc))
                    SetStale(cc, stale: false);
                continue;
            }

            // New content and title: not greyed. Equations and inline errors are drawn together at
            // the end. Not hidden expressions: a copy of hidden text comes out empty.
            if (tag.Kind == ControlKind.Inline && result.Kind != ResultKind.Hidden)
                inline.Add(new InlineResult(cc, result, fingerprint));
            else
            {
                var start = System.Diagnostics.Stopwatch.GetTimestamp();
                Render(cc, tag.Kind, result, fingerprint, s.TextWidth);
                blockTicks += System.Diagnostics.Stopwatch.GetTimestamp() - start;
            }
            ++redrawn;
        }
        var scanned = timer.Elapsed;
        WriteInline(s.Doc.Application, inline);
        if (s.Only is null)
            SetOutOfDate(s.Doc, false);
        else if (s.Limit < s.Items.Count)
            SetOutOfDate(s.Doc, true); // what follows may depend on it

        var total = s.Only?.Count ?? controls.Count;
        var fields = s.Inputs.Count;
        Log.Info($"{(s.Only is null ? "Recalculated" : "Calculated what was edited in")} '{s.Doc.Name}': {total} items, {fields} inputs, {errors} with errors, {redrawn} redrawn " +
                 $"in {timer.Elapsed.TotalSeconds:F1} s: inputs {inputsDone.TotalSeconds:F1} s, blocks {Seconds(blockTicks):F1} s, " +
                 $"rest of the scan {(scanned - inputsDone).TotalSeconds - Seconds(blockTicks):F1} s, {inline.Count} equations {(timer.Elapsed - scanned).TotalSeconds:F1} s " +
                 $"({Seconds(_insertTicks):F1} s in InsertXML, {Seconds(_deleteTicks):F1} s deleting tables)");
        if (total == 0 && fields == 0)
            return "Calcpad: nothing to calculate. Select text and press Convert to Calcpad (Alt+Shift+Enter), type {= expression } in the text, or add a Calcpad Block.";

        if (s.Only is not null)
            return $"Calcpad: calculated {(total == 1 ? "the expression" : $"{total} expressions")}" +
                   (firstError is null ? "" : $", {errors} error{(errors == 1 ? "" : "s")}: {firstError}") +
                   (IsOutOfDate(s.Doc) ? ". Other results are out of date: Calculate (Ctrl+Alt+F9) updates them." : ".");

        return $"Calcpad: {total} item{(total == 1 ? "" : "s")}" +
               (fields == 0 ? "" : $", {fields} input{(fields == 1 ? "" : "s")}") +
               $", {errors} error{(errors == 1 ? "" : "s")}" +
               (firstError is null ? "" : $". First error: {firstError}  (Ctrl+Alt+F8 goes to it)");
    }

    /// <summary>Light red behind an input field whose value can't be used; removed once it can. Changes nothing otherwise.</summary>
    private static void MarkInput(Word.ContentControl cc, bool error)
    {
        var shading = cc.Range.Font.Shading;
        var marked = shading.BackgroundPatternColor == InputErrorColor;
        if (error != marked)
            shading.BackgroundPatternColor = error ? InputErrorColor : Word.WdColor.wdColorAutomatic;
    }

    /// <summary>#FFC7CE (Word colours are BGR).</summary>
    private const Word.WdColor InputErrorColor = (Word.WdColor)0xCEC7FF;

    /// <summary>What an input field holds: its text, or for a list the chosen item's value (its text if it has none).</summary>
    private static string ReadInput(Word.ContentControl cc)
    {
        if (cc.ShowingPlaceholderText)
            return string.Empty;

        var text = cc.Range.Text ?? string.Empty;
        if (cc.Type is Word.WdContentControlType.wdContentControlDropdownList or Word.WdContentControlType.wdContentControlComboBox)
        {
            var entries = cc.DropdownListEntries;
            for (int i = 1, n = entries.Count; i <= n; ++i)
            {
                var entry = entries[i];
                if (entry.Text == text)
                    return string.IsNullOrEmpty(entry.Value) ? entry.Text : entry.Value;
            }
        }
        return text;
    }

    private sealed record InlineResult(Word.ContentControl Control, EvaluationResult Result, string Fingerprint);

    /// <summary>From this many equations on, they are drawn together; fewer are inserted one by one.</summary>
    private const int DrawTogetherFrom = 6;

    /// <summary>True while the scratch document is being made: <see cref="WordEvents"/> doesn't watch it.</summary>
    public static bool MakingScratch { get; private set; }

    /// <summary>
    /// Draws equations and inline errors. Many at once go through a hidden scratch document:
    /// one InsertXML of all their paragraphs, then a FormattedText copy into each control, which
    /// takes the document's own formatting. In a long document an InsertXML costs about 0.1 s
    /// and a copy about 15 ms (the v2 beam, 2026-11-18). The scratch document is closed at the
    /// end: an invisible document left open could keep Word running after its last window closes.
    /// If anything fails, the rest are inserted one by one.
    /// </summary>
    private static void WriteInline(Word.Application app, List<InlineResult> results)
    {
        var drawn = 0;
        if (results.Count >= DrawTogetherFrom)
        {
            Word.Document? scratch = null;
            try
            {
                MakingScratch = true;
                try
                {
                    scratch = app.Documents.Add(Visible: false);
                }
                finally
                {
                    MakingScratch = false;
                }

                var start = System.Diagnostics.Stopwatch.GetTimestamp();
                scratch.Content.InsertXML(WordMarkup.Package(string.Concat(results.Select(r => InlineParagraph(r.Result)))));
                var inserted = System.Diagnostics.Stopwatch.GetTimestamp();
                _insertTicks += inserted - start;
                var paragraphs = scratch.Paragraphs;
                if (paragraphs.Count != results.Count)
                    throw new InvalidOperationException($"the scratch document has {paragraphs.Count} paragraphs for {results.Count} results");

                long copyTicks = 0;
                var paragraph = paragraphs.First;
                foreach (var r in results)
                {
                    var content = paragraph.Range;
                    content.MoveEnd(Word.WdUnits.wdCharacter, -1); // not the paragraph mark
                    var copy = System.Diagnostics.Stopwatch.GetTimestamp();
                    r.Control.LockContents = false;
                    r.Control.Range.FormattedText = content.FormattedText;
                    copyTicks += System.Diagnostics.Stopwatch.GetTimestamp() - copy;
                    Label(r.Control, r.Result, r.Fingerprint);
                    ++drawn;
                    paragraph = paragraph.Next();
                }
                Log.Info($"Equations together: InsertXML {Seconds(inserted - start):F2} s, copies {Seconds(copyTicks):F2} s, " +
                         $"labels and the rest {Seconds(System.Diagnostics.Stopwatch.GetTimestamp() - inserted - copyTicks):F2} s");
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException)
            {
                Log.Error($"Drawing {results.Count} equations together failed after {drawn}: drawing the rest one by one", ex);
            }
            finally
            {
                try { scratch?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch (COMException ex) { Log.Error("Closing the scratch document failed", ex); }
            }
        }

        foreach (var r in results.Skip(drawn))
        {
            Write(r.Control, WordMarkup.Package(InlineParagraph(r.Result)));
            Label(r.Control, r.Result, r.Fingerprint);
        }
    }

    /// <summary>An inline result's one paragraph: its equation, or for a hidden expression or an error, its source.</summary>
    private static string InlineParagraph(EvaluationResult result) => result.Kind switch
    {
        ResultKind.Math => WordMarkup.MathParagraph(result.Omml),
        ResultKind.Hidden => WordMarkup.TextParagraph(result.Source, TextLook.Hidden),
        _ => WordMarkup.TextParagraph(result.Source, TextLook.Error),
    };

    /// <summary>
    /// The title and tag of an inline result just drawn. Results are read-only (Edit Source or a
    /// double-click changes them); an error stays editable source, with the message as its title.
    /// </summary>
    private static void Label(Word.ContentControl cc, EvaluationResult result, string fingerprint)
    {
        if (result.Kind == ResultKind.Error)
        {
            cc.Title = Clip(ErrorTitlePrefix + result.Error);
            cc.Tag = new ControlTag(ControlState.Error, fingerprint, result.Source, ControlKind.Inline).ToString();
            return;
        }
        cc.Title = Clip((result.Kind == ResultKind.Hidden ? "Calcpad (hidden): " : "Calcpad: ") + result.Source);
        cc.Tag = new ControlTag(ControlState.Result, fingerprint, result.Source, ControlKind.Inline).ToString();
        cc.LockContents = true;
    }

    private static void Render(Word.ContentControl cc, ControlKind kind, EvaluationResult result, string fingerprint, double? textWidth)
    {
        switch (result.Kind)
        {
            case ResultKind.Math or ResultKind.Hidden:
                Write(cc, WordMarkup.Package(InlineParagraph(result)));
                Label(cc, result, fingerprint);
                break;

            case ResultKind.Block:
                Write(cc, result.Block!.WordPackageFor(textWidth), block: true); // made on the background thread
                cc.Title = result.ErrorCount == 0
                    ? Clip("Calcpad block: " + FirstLine(result.Source))
                    : Clip($"{ErrorTitlePrefix}{result.Error} ({result.ErrorCount} error{(result.ErrorCount == 1 ? "" : "s")} in this block)");
                cc.Tag = new ControlTag(ControlState.Result, fingerprint, result.Source, kind).ToString();
                cc.LockContents = true;
                break;

            default:
                // Errors stay as editable source so they can be fixed in place; the message is the title.
                ShowSource(cc, result.Source, TextLook.Error, ErrorTitlePrefix + result.Error, ControlState.Error, fingerprint);
                break;
        }
    }

    private static void ShowSource(Word.ContentControl cc, string source, TextLook look, string title, ControlState state, string fingerprint)
    {
        var kind = ControlTag.Parse(cc.Tag)?.Kind ?? ControlKind.Inline;
        var body = kind == ControlKind.Block
            ? WordMarkup.SourceParagraphs(source, look)
            : WordMarkup.TextParagraph(source, look);
        Write(cc, WordMarkup.Package(body), block: kind == ControlKind.Block);
        cc.Title = Clip(title);
        cc.Tag = new ControlTag(state, fingerprint, source, kind).ToString();
    }

    /// <summary>
    /// Replaces the control's content. Always via InsertXML: see <see cref="WordMarkup"/>.
    /// Only a block can hold tables, which have to go first; inline results skip that lookup.
    /// </summary>
    private static void Write(Word.ContentControl cc, string package, bool block = false)
    {
        cc.LockContents = false;
        var range = cc.Range;
        if (block)
        {
            var delete = System.Diagnostics.Stopwatch.GetTimestamp();
            DeleteTables(range);
            _deleteTicks += System.Diagnostics.Stopwatch.GetTimestamp() - delete;
        }
        var start = System.Diagnostics.Stopwatch.GetTimestamp();
        range.InsertXML(package);
        _insertTicks += System.Diagnostics.Stopwatch.GetTimestamp() - start;
    }

    /// <summary>Time spent in InsertXML and deleting tables during the current <see cref="Draw"/> (Word's thread only), for the log.</summary>
    private static long _insertTicks, _deleteTicks;

    private static double Seconds(long ticks) => (double)ticks / System.Diagnostics.Stopwatch.Frequency;

    // ---- Out-of-date results ---------------------------------------------------------------
    //
    // While a calculation runs in the background (and after it is cancelled), results that may
    // change are greyed: a light grey background, washed-out pictures, and "Out of date: " in
    // front of the title. All three can be undone exactly, which matters for results that turn
    // out unchanged and are not redrawn. The title also keeps the mark in a saved document.

    private const string StaleTitlePrefix = "Out of date: ";

    private static bool IsStale(Word.ContentControl cc) =>
        (cc.Title ?? string.Empty).StartsWith(StaleTitlePrefix, StringComparison.Ordinal);

    /// <summary>The control's title without the out-of-date mark.</summary>
    private static string TitleOf(Word.ContentControl cc)
    {
        var title = cc.Title ?? string.Empty;
        return title.StartsWith(StaleTitlePrefix, StringComparison.Ordinal) ? title[StaleTitlePrefix.Length..] : title;
    }

    /// <summary>
    /// The results that the calculation now running may change: everything from the first item
    /// that is new or being edited (the worksheet runs top to bottom, so nothing before it
    /// changes), or everything if none is (Recalculate, a settings change). Reads only tags.
    /// </summary>
    private static List<Word.ContentControl> StaleCandidates(Snapshot s)
    {
        if (s.Only is not null)
            return []; // drawing only what was edited, which shows its source meanwhile: nothing else changes

        // The first item (in worksheet order) that is new, edited, or an input field with a new value.
        var first = s.Parsed.FindIndex(p => p.ShowsSource || p.Fingerprint is "" or "-");
        var firstItem = first >= 0 ? s.ItemOf[first] : int.MaxValue;
        foreach (var input in s.Inputs)
            if (input.Changed)
                firstItem = Math.Min(firstItem, input.Item);
        if (firstItem == int.MaxValue)
            firstItem = 0;

        var candidates = new List<Word.ContentControl>();
        for (var i = 0; i < s.Controls.Count; ++i)
        {
            // Only results already drawn: new items still show their source, edited ones too.
            var tag = s.Parsed[i];
            if (s.ItemOf[i] >= firstItem && tag.State == ControlState.Result && tag.Fingerprint is not ("" or "-"))
                candidates.Add(s.Controls[i]);
        }
        return candidates;
    }

    private static void SetStale(Word.ContentControl cc, bool stale)
    {
        var title = TitleOf(cc);
        var locked = cc.LockContents;
        cc.LockContents = false;
        try
        {
            var range = cc.Range;
            range.Font.Shading.BackgroundPatternColor = stale ? Word.WdColor.wdColorGray10 : Word.WdColor.wdColorAutomatic;
            // Pictures in grey, not washed out (watermark): a drawing that faded almost to white
            // and came back flashed.
            var pictures = range.InlineShapes;
            for (var i = 1; i <= pictures.Count; ++i)
            {
                pictures[i].PictureFormat.ColorType = stale
                    ? Microsoft.Office.Core.MsoPictureColorType.msoPictureGrayscale
                    : Microsoft.Office.Core.MsoPictureColorType.msoPictureAutomatic;
            }
            cc.Title = stale ? Clip(StaleTitlePrefix + title) : title;
        }
        finally
        {
            cc.LockContents = locked;
        }
    }

    /// <summary>
    /// Word refuses to replace a range that contains a whole table ("The range cannot be
    /// deleted"), as a block's output can. Deleting the tables first is allowed.
    /// </summary>
    private static void DeleteTables(Word.Range range)
    {
        var tables = range.Tables;
        for (var i = tables.Count; i >= 1; --i)
            tables[i].Delete();
    }

    private static string ReadEditedSource(Word.ContentControl cc, ControlKind kind)
    {
        if (cc.ShowingPlaceholderText)
            return string.Empty;

        return kind == ControlKind.Block
            ? SourceNormalizer.NormalizeBlock(cc.Range.Text)
            : SourceNormalizer.Normalize(cc.Range.Text);
    }

    private static void PlaceCursorAtEnd(Word.ContentControl cc)
    {
        var end = cc.Range;
        end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        end.Select();
    }

    // ---- Turning text into controls -----------------------------------------------------

    /// <summary>
    /// Shift+Enter while editing an expression adds a line (a Word line break); so does a
    /// paragraph break pasted or typed into it. Each line becomes its own expression: the
    /// control is unwrapped (its text stays) and every non-empty line is wrapped again, in place.
    /// The source the cursor is in is left alone until the user leaves it.
    /// </summary>
    private static void SplitMultiLineExpressions(Word.Document doc, string? activeId)
    {
        foreach (var cc in CalcpadControls(doc))
        {
            var tag = ControlTag.Parse(cc.Tag)!;
            if (tag.Kind != ControlKind.Inline || !tag.ShowsSource || cc.ID == activeId)
                continue;

            var content = cc.Range;
            var text = content.Text ?? string.Empty;
            if (!text.Contains('\r') && !text.Contains('\v'))
                continue;

            // A range around the control (its boundaries included) follows the text once the control is gone.
            var lines = Slice(content, content.Start - 1, content.End + 1);
            cc.LockContents = false;
            cc.Delete(false);
            WrapLines(lines);
        }
    }

    /// <summary>Wraps each line of <paramref name="range"/> (split at paragraph marks and line breaks) as an expression.</summary>
    private static void WrapLines(Word.Range range)
    {
        var text = range.Text ?? string.Empty;
        if (text.Length != range.End - range.Start)
        {
            // Something in the text takes up positions differently (a field, an object); go by paragraphs.
            WrapAsExpressions(range);
            return;
        }

        var lines = new List<(int Start, int End)>();
        var lineStart = 0;
        for (var i = 0; i <= text.Length; ++i)
        {
            if (i < text.Length && text[i] is not ('\r' or '\v' or '\a'))
                continue;
            if (i > lineStart)
                lines.Add((range.Start + lineStart, range.Start + i));
            lineStart = i + 1;
        }

        // Back to front, so earlier positions stay valid.
        for (var i = lines.Count - 1; i >= 0; --i)
        {
            var piece = Slice(range, lines[i].Start, lines[i].End);
            Trim(piece);
            if (piece.End <= piece.Start || piece.ContentControls.Count > 0)
                continue;

            var source = piece.Text ?? string.Empty;
            if (source.Contains(DelimiterScanner.Open, StringComparison.Ordinal))
                continue; // the delimiter pass handles {= ... }

            CreateInline(piece, SourceNormalizer.Normalize(source));
        }
    }

    private static void WrapAsExpressions(Word.Range selection)
    {
        var scope = selection.Start == selection.End ? selection.Paragraphs.First.Range : selection;
        var pieces = new List<Word.Range>();
        var paragraphs = scope.Paragraphs;
        for (int i = 1, n = paragraphs.Count; i <= n; ++i)
        {
            var p = paragraphs[i].Range;
            var piece = Slice(scope, Math.Max(p.Start, scope.Start), Math.Min(p.End, scope.End));
            Trim(piece);
            if (piece.End <= piece.Start || piece.ContentControls.Count > 0)
                continue;

            var text = piece.Text ?? string.Empty;
            if (text.Contains(DelimiterScanner.Open, StringComparison.Ordinal))
                continue; // the delimiter pass handles {= ... } inside prose

            pieces.Add(piece);
        }

        // Back to front, so earlier positions stay valid.
        for (var i = pieces.Count - 1; i >= 0; --i)
            CreateInline(pieces[i], SourceNormalizer.Normalize(pieces[i].Text));
    }

    /// <summary>
    /// Finds every <c>{= … }</c> in one story and every fenced block (a paragraph that is only
    /// <c>{=</c>, the script, and a paragraph that is only <c>}</c>), and turns each into a control.
    /// </summary>
    private static void ConvertDelimiters(Word.Range story)
    {
        var starts = new List<int>();
        var range = story.Duplicate;
        var find = range.Find;
        find.ClearFormatting();
        find.Text = DelimiterScanner.Open;
        find.Forward = true;
        find.Wrap = Word.WdFindWrap.wdFindStop;
        find.Format = false;
        find.MatchWildcards = false;
        find.MatchCase = false;
        while (starts.Count < MaxDelimiters && find.Execute())
        {
            starts.Add(range.Start);
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        // Plan first (front to back, so text inside a fenced block is not taken as inline), apply back to front.
        var conversions = new List<Action>();
        var skipUntil = -1;
        foreach (var start in starts)
        {
            if (start < skipUntil)
                continue;

            var at = Slice(story, start, start);
            if (IsInsideCalcpadControl(at))
                continue;

            var paragraph = at.Paragraphs.First;
            if (BlockFence.IsOpening(paragraph.Range.Text))
            {
                var following = new List<Word.Paragraph>();
                var close = BlockFence.FindClose(Following(paragraph, following));
                if (close < 0)
                    continue; // not closed: leave the text alone

                var open = paragraph.Range;
                var closing = following[close].Range;
                var inner = close > 0 ? Slice(story, following[0].Range.Start, following[close - 1].Range.End) : null;
                skipUntil = closing.End;
                conversions.Add(() =>
                {
                    // Word ranges follow edits, so 'inner' stays on the script while the fences go.
                    closing.Delete();
                    open.Delete();
                    if (inner is not null)
                        CreateBlock(inner, SourceNormalizer.NormalizeBlock(inner.Text));
                });
                continue;
            }

            var tail = Slice(story, start, paragraph.Range.End).Text ?? string.Empty;
            var length = DelimiterScanner.MatchLength(tail, 0);
            if (length < 0)
                continue; // not closed in this paragraph: leave the text alone

            var target = Slice(story, start, start + length);
            if (target.Text != tail[..length])
            {
                // Fields, controls or other objects inside shift Word's positions; don't guess.
                Log.Info($"Skipped {{= ... }} at {start} in story {story.StoryType}: its text contains objects.");
                continue;
            }
            var source = DelimiterScanner.Inner(tail[..length]);
            conversions.Add(() => CreateInline(target, source));
        }

        for (var i = conversions.Count - 1; i >= 0; --i)
            conversions[i]();

        static IEnumerable<string> Following(Word.Paragraph paragraph, List<Word.Paragraph> visited)
        {
            for (var p = paragraph.Next(); p is not null; p = p.Next())
            {
                visited.Add(p);
                yield return p.Range.Text ?? string.Empty;
            }
        }
    }

    private static void CreateInline(Word.Range range, string source)
    {
        var cc = range.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, range);
        // Result state with no fingerprint: the next pass evaluates and draws it.
        cc.Tag = new ControlTag(ControlState.Result, string.Empty, source).ToString();
        cc.Title = Clip("Calcpad: " + source);
    }

    /// <summary>Replaces whole paragraphs with a converted script, then gives headings Word's Heading styles.</summary>
    private static void InsertConverted(Word.Range paragraphs, ConversionResult conversion)
    {
        var story = paragraphs.Duplicate;
        story.WholeStory();
        var storyLength = story.End;
        DeleteTables(paragraphs); // a converted block's output may hold one
        var start = paragraphs.Start;
        var replaced = paragraphs.End - paragraphs.Start;

        // Replacing whole paragraphs, InsertXML drops the last paragraph mark it inserts and joins
        // that paragraph to the text that follows. An empty last paragraph takes that hit instead.
        paragraphs.InsertXML(WordMarkup.Package(WordMarkup.ConvertedContent(conversion.Parts) + "<w:p/>"));

        story.WholeStory();
        var inserted = Slice(story, start, Math.Min(story.End, start + replaced + story.End - storyLength));
        var all = inserted.Paragraphs;
        for (int i = 1, n = all.Count; i <= n; ++i)
        {
            var p = all[i];
            var heading = p.OutlineLevel switch
            {
                Word.WdOutlineLevel.wdOutlineLevel1 => Word.WdBuiltinStyle.wdStyleHeading1,
                Word.WdOutlineLevel.wdOutlineLevel2 => Word.WdBuiltinStyle.wdStyleHeading2,
                Word.WdOutlineLevel.wdOutlineLevel3 => Word.WdBuiltinStyle.wdStyleHeading3,
                Word.WdOutlineLevel.wdOutlineLevel4 => Word.WdBuiltinStyle.wdStyleHeading4,
                _ => (Word.WdBuiltinStyle?)null,
            };
            if (heading is null || p.Range.ParentContentControl is not null)
                continue;

            p.Reset(); // drop the outline level that marked it; the style brings its own
            p.Range.set_Style(heading.Value);
        }
    }

    /// <summary>Wraps whole paragraphs in a block control.</summary>
    private static Word.ContentControl CreateBlock(Word.Range paragraphs, string source)
    {
        var range = WholeParagraphs(paragraphs);
        var story = range.Duplicate;
        story.WholeStory();
        if (range.End >= story.End)
        {
            // Keep an ordinary paragraph after the block when it reaches the end of the text.
            story.InsertParagraphAfter();
        }
        var cc = range.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText, range);
        cc.Tag = new ControlTag(ControlState.Result, string.Empty, source, ControlKind.Block).ToString();
        cc.Title = Clip("Calcpad block: " + FirstLine(source));
        return cc;
    }

    // ---- Finding stories and controls ---------------------------------------------------

    private static List<Word.Range> Stories(Word.Document doc)
    {
        var stories = new List<Word.Range>();
        var all = doc.StoryRanges;
        foreach (var type in StoryOrder)
        {
            Word.Range? story;
            try
            {
                story = all[type];
            }
            catch (COMException)
            {
                continue; // this document has no such story
            }
            for (; story is not null; story = story.NextStoryRange)
                stories.Add(story);
        }
        return stories;
    }

    /// <summary>All Calcpad expressions and blocks, in evaluation order. Linked headers are visited once.</summary>
    private static List<Word.ContentControl> CalcpadControls(Word.Document doc) => Controls(doc, ControlTag.IsCalcpad);

    /// <summary>All input fields, in evaluation order.</summary>
    private static List<Word.ContentControl> InputControls(Word.Document doc) => Controls(doc, InputField.IsInput);

    /// <summary>Everything the worksheet runs: Calcpad expressions, blocks and input fields, in evaluation order.</summary>
    private static List<Word.ContentControl> WorksheetControls(Word.Document doc) =>
        Controls(doc, tag => ControlTag.IsCalcpad(tag) || InputField.IsInput(tag));

    private static List<Word.ContentControl> Controls(Word.Document doc, Func<string?, bool> wanted)
    {
        var result = new List<Word.ContentControl>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var story in Stories(doc))
        {
            var found = new List<(int Start, Word.ContentControl Control)>();
            var controls = story.ContentControls;
            for (int i = 1, n = controls.Count; i <= n; ++i)
            {
                var cc = controls[i];
                if (wanted(cc.Tag) && seen.Add(cc.ID))
                    found.Add((cc.Range.Start, cc));
            }
            result.AddRange(found.OrderBy(f => f.Start).Select(f => f.Control));
        }
        return result;
    }

    private static Word.ContentControl? FindCalcpadControl(Word.Range range)
    {
        for (var cc = range.ParentContentControl; cc is not null; cc = cc.ParentContentControl)
            if (ControlTag.IsCalcpad(cc.Tag))
                return cc;

        // The whole control is selected.
        var inside = range.ContentControls;
        if (inside.Count == 1 && ControlTag.IsCalcpad(inside[1].Tag))
            return inside[1];

        return null;
    }

    private static bool IsInsideCalcpadControl(Word.Range range)
    {
        for (var cc = range.ParentContentControl; cc is not null; cc = cc.ParentContentControl)
            if (ControlTag.IsCalcpad(cc.Tag))
                return true;

        return false;
    }

    // ---- Helpers ------------------------------------------------------------------------

    private Word.Document ActiveDocument() =>
        TryActiveDocument() ?? throw new UserMessageException("Open a document first. (Calcpad can't change documents in Protected View.)");

    private Word.Document? TryActiveDocument()
    {
        try
        {
            return app.Documents.Count == 0 ? null : app.ActiveDocument;
        }
        catch (COMException)
        {
            return null; // e.g. Protected View is active
        }
    }

    private static string? DocumentFolder(Word.Document doc)
    {
        try { return string.IsNullOrEmpty(doc.Path) ? null : doc.Path; }
        catch (COMException) { return null; }
    }

    private static DocumentSettings LoadSettings(Word.Document doc)
    {
        DocumentSettings settings;
        try
        {
            settings = DocumentSettings.Parse(doc.Variables[SettingsVariable].Value);
        }
        catch (COMException)
        {
            settings = new DocumentSettings(); // never set in this document
        }
        // Word's own math font setting (Equation Options) is the one that counts.
        return settings with { MathFont = MathFontOf(doc) };
    }

    private static string MathFontOf(Word.Document doc)
    {
        try { return string.IsNullOrWhiteSpace(doc.OMathFontName) ? MathFonts.Default : doc.OMathFontName; }
        catch (COMException) { return MathFonts.Default; }
    }

    private static bool IsOpen(Word.Document doc)
    {
        try { _ = doc.Name; return true; }
        catch (COMException) { return false; }
    }

    /// <summary>The control is still in its document (Word throws for one that was deleted).</summary>
    private static bool IsAlive(Word.ContentControl cc)
    {
        try { _ = cc.ID; return true; }
        catch (COMException) { return false; }
    }

    // ---- Math font ----------------------------------------------------------------------

    private volatile Task? _fontInstall;
    private volatile CancellationTokenSource? _fontInstallCancel;
    private readonly HashSet<string> _installedThisSession = new(StringComparer.OrdinalIgnoreCase);
    private HashSet<string>? _wordFonts;

    private (string Current, long At, IReadOnlyList<(MathFontInfo Font, bool Installed)> Choices)? _fontChoices;

    /// <summary>
    /// The fonts the ribbon offers, alphabetically, with whether each is installed: the starters,
    /// the other math fonts found on this computer, and the document's own font. Kept for a moment,
    /// because the ribbon asks once per item.
    /// </summary>
    public IReadOnlyList<(MathFontInfo Font, bool Installed)> MathFontChoices()
    {
        var current = CurrentMathFont();
        var now = Environment.TickCount64;
        if (_fontChoices is { } cached && cached.Current == current && now - cached.At < 2000)
            return cached.Choices;

        var installed = InstalledFonts();
        var math = new HashSet<string>(InstalledMathFonts.Families(), StringComparer.OrdinalIgnoreCase);
        math.UnionWith(_installedThisSession);
        var choices = MathFonts.Choices(math, installed.Contains, current);
        _fontChoices = (current, now, choices);
        return choices;
    }

    /// <summary>Finds the installed math fonts in the background at start-up, so the ribbon's list is ready.</summary>
    public static void DiscoverMathFonts() => Task.Run(() =>
    {
        try { InstalledMathFonts.Families(); }
        catch (Exception ex) { Log.Error("Looking for math fonts failed", ex); }
    });

    public string CurrentMathFont() => TryActiveDocument() is { } doc ? MathFontOf(doc) : MathFonts.Default;

    /// <summary>
    /// Sets the document's math font and redraws the equations in it. A free font that isn't
    /// installed is downloaded and installed for the user first (in the background), but only
    /// with <paramref name="allowDownload"/>; otherwise <see cref="FontNotInstalledException"/> lets
    /// the caller ask.
    /// </summary>
    public string SetMathFont(string family, bool allowDownload)
    {
        var doc = ActiveDocument();
        var font = MathFontChoices().FirstOrDefault(c => c.Font.Family.Equals(family, StringComparison.OrdinalIgnoreCase));
        if (font.Font is null)
            throw new UserMessageException($"{family} is not one of the math fonts Calcpad offers.");

        if (font.Installed)
            return ApplyMathFont(doc, font.Font.Family);

        if (font.Font.DownloadUrl is null)
            throw new UserMessageException($"{font.Font.Family} is not installed on this computer. {font.Font.Note} Then restart Word and choose it again.");
        if (!allowDownload)
            throw new FontNotInstalledException(font.Font);
        if (_fontInstall is not null)
            throw new UserMessageException("A font is already being installed. Try again in a moment.");

        var cancel = new CancellationTokenSource();
        var install = Task.Run(() => FontInstaller.InstallAsync(font.Font, cancel.Token));
        _fontInstallCancel = cancel;
        _fontInstall = install;
        UiThread.WhenDone(install, () =>
        {
            string status;
            try
            {
                if (install.IsCanceled || cancel.IsCancellationRequested)
                {
                    Log.Info($"Installing {font.Font.Family} was cancelled");
                    status = LastStatus = $"Calcpad: installing {font.Font.Family} was cancelled.";
                }
                else if (install.IsFaulted)
                {
                    Log.Error($"Installing {font.Font.Family} failed", install.Exception);
                    status = LastStatus = $"Calcpad: {font.Font.Family} could not be installed: {install.Exception?.GetBaseException().Message}";
                }
                else
                {
                    _installedThisSession.Add(font.Font.Family);
                    _fontChoices = null;
                    status = IsOpen(doc) ? ApplyMathFont(doc, font.Font.Family) : LastStatus = $"Calcpad: {font.Font.Family} is installed.";
                }
            }
            finally
            {
                _fontInstall = null; // only now: a calculation started above already counts as working
                _fontInstallCancel = null;
                cancel.Dispose();
            }
            BackgroundFinished?.Invoke(status);
        });
        return $"Calcpad: downloading and installing {font.Font.Family} (free; {font.Font.Note})…";
    }

    private string ApplyMathFont(Word.Document doc, string family)
    {
        using var undo = new UndoScope(app, "Calcpad Math Font");
        doc.OMathFontName = family;
        var status = RecalculateCore(doc, keepActiveSource: false);
        return status.StartsWith("Calcpad: calculating", StringComparison.Ordinal) ? status : $"{status}. Math font: {family}.";
    }

    private HashSet<string> InstalledFonts()
    {
        if (_wordFonts is null)
        {
            _wordFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var names = app.FontNames;
                for (int i = 1, n = names.Count; i <= n; ++i)
                    _wordFonts.Add(names[i]);
            }
            catch (COMException ex)
            {
                Log.Error("Reading Word's font list failed", ex);
            }
        }
        // Word reads its font list at start-up; fonts installed since are registered with Windows.
        var all = new HashSet<string>(_wordFonts, StringComparer.OrdinalIgnoreCase);
        all.UnionWith(FontInstaller.RegisteredFamilies());
        all.UnionWith(_installedThisSession);
        return all;
    }

    private static void SaveSettings(Word.Document doc, DocumentSettings settings)
    {
        var value = settings.Serialize();
        try
        {
            doc.Variables[SettingsVariable].Value = value;
        }
        catch (COMException)
        {
            doc.Variables.Add(SettingsVariable, value);
        }
    }

    /// <summary>The paragraphs touched by <paramref name="range"/>, from the start of the first to the end of the last.</summary>
    private static Word.Range WholeParagraphs(Word.Range range)
    {
        var paragraphs = range.Paragraphs;
        return Slice(range, paragraphs.First.Range.Start, paragraphs.Last.Range.End);
    }

    private static IEnumerable<string> ParagraphTexts(Word.Range range)
    {
        var paragraphs = range.Paragraphs;
        for (int i = 1, n = paragraphs.Count; i <= n; ++i)
            yield return paragraphs[i].Range.Text ?? string.Empty;
    }

    /// <summary>A range over [start, end) in the same story as <paramref name="within"/>.</summary>
    private static Word.Range Slice(Word.Range within, int start, int end)
    {
        var r = within.Duplicate;
        r.SetRange(start, end);
        return r;
    }

    private static void Trim(Word.Range range)
    {
        range.MoveEndWhile(" \t\r\n\a\v\f\u00A0", Word.WdConstants.wdBackward);
        if (range.End > range.Start)
            range.MoveStartWhile(" \t\u00A0", Word.WdConstants.wdForward);
    }

    private static string FirstLine(string source)
    {
        var line = source.Split('\n').FirstOrDefault(l => l.Trim().Length > 0)?.Trim() ?? string.Empty;
        return line.Length <= 60 ? line : line[..59] + "…";
    }

    private static string Clip(string s) => s.Length <= MaxTitle ? s : s[..(MaxTitle - 1)] + "…";
}
