using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using CalcpadWord.AddIn.Interop;
using CalcpadWord.Engine;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CalcpadWord.AddIn;

/// <summary>
/// The COM add-in Word loads. It wires the ribbon, keyboard shortcuts and Word events to
/// <see cref="DocumentController"/>, and is exposed as
/// <c>Application.COMAddIns("CalcpadWord.Connect").Object</c> for macros and tests.
/// </summary>
[ComVisible(true)]
[Guid(ClassId)]
[ProgId(ProgIdName)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
{
    public const string ClassId = AddInIdentity.ClassId;
    public const string ProgIdName = AddInIdentity.ProgId;
    private const int MaxErrorMenuItems = 30;

    // Word's right-click menus that can open on a Calcpad expression (ids checked with
    // CommandBars.GetLabelMso): ContextMenuEquation on a result; the text menus on source
    // text, in tables, lists, headings, headers and footers.
    private static readonly string[] ContextMenus =
    [
        "ContextMenuEquation", "ContextMenuText", "ContextMenuTextLinked", "ContextMenuTable", "ContextMenuList",
        "ContextMenuHeading", "ContextMenuHeaderArea", "ContextMenuFooterArea", "ContextMenuSpell", "ContextMenuGrammar",
        "ContextMenuInlinePicture", // plots in a block's output
    ];

    private static readonly string[] ContextItemIds =
        [.. ContextMenus.SelectMany(menu => new[] { $"CalcpadEditSource_{menu}", $"CalcpadCalculate_{menu}", $"CalcpadBlockToText_{menu}" })];

    private Word.Application? _app;
    private DocumentController? _controller;
    private IRibbonUI? _ribbon;
    private WordEvents? _events;
    private KeyboardShortcuts? _keys;

    // Whether Enter at the cursor belongs to Calcpad. Kept current from Word events and after every
    // command, because the keyboard hook must decide without calling Word.
    private volatile bool _enterLeavesControl;

    private DocumentController Controller =>
        _controller ?? throw new InvalidOperationException("The Calcpad add-in is not connected to Word.");

    // ---- IDTExtensibility2 --------------------------------------------------------------

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            Log.Info($"Connecting (mode {ConnectMode}), version {Version}, .NET {Environment.Version}");
            UiThread.Initialize();
            _app = (Word.Application)Application;
            _controller = new DocumentController(_app);
            _controller.WarmUp();
            ((COMAddIn)AddInInst).Object = this;

            _controller.BackgroundFinished += OnBackgroundFinished;
            _events = new WordEvents(_app, OnControlExited, OnDoubleClick, OnRightClick, RefreshCursorState, OnDocumentChange);
            _events.Start();
            _keys = new KeyboardShortcuts(
            [
                new(KeyboardShortcuts.Evaluate, () => Run("Convert", () => Controller.EvaluateOrEditSource())),
                new(KeyboardShortcuts.Recalculate, () => Run("Calculate", () => Controller.Recalculate())),
                new(KeyboardShortcuts.NextError, () => Run("Next Error", () => Controller.NextError())),
                // Enter in an expression finishes it; Shift+Enter (Word's line break) adds another.
                new(KeyboardShortcuts.Enter, () => Run("Enter", () => { PressEnterCore(); return string.Empty; }), Claims: () => _enterLeavesControl),
                // Esc stops a background calculation; otherwise it is Word's.
                new(KeyboardShortcuts.Escape, () => Run("Cancel", () => Controller.Cancel()), Claims: () => _controller?.CanCancel == true),
            ]);
            DocumentController.DiscoverMathFonts();
            RefreshCursorState();
        }
        catch (Exception ex)
        {
            Log.Error("OnConnection failed", ex);
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        Log.Info($"Disconnecting ({RemoveMode})");
        _keys?.Dispose();
        _events?.Dispose();
        _controller?.CancelBackgroundWork();
        UiThread.Shutdown();
        _keys = null;
        _events = null;
        _controller = null;
        _ribbon = null;
        _app = null;
    }

    public void OnAddInsUpdate(ref Array custom) { }

    public void OnStartupComplete(ref Array custom) { }

    public void OnBeginShutdown(ref Array custom) { }

    // ---- IRibbonExtensibility -----------------------------------------------------------

    public string GetCustomUI(string RibbonID)
    {
        using var stream = typeof(Connect).Assembly.GetManifestResourceStream("CalcpadWord.AddIn.Ribbon.xml")!;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd()
            .Replace("<!--SYMBOL_GALLERIES-->", SymbolPalette.RibbonXml(), StringComparison.Ordinal)
            .Replace("<!--CONTEXT_MENUS-->", ContextMenusXml(), StringComparison.Ordinal);
    }

    // ---- Symbols: a grid of cells scrolled by rows, as Word's equation Symbols gallery ------

    private int _symbolTopRow;

    /// <summary>Pictures named in the ribbon markup (customUI loadImage), drawn when the ribbon asks.</summary>
    public object? LoadImage(string imageId) => SafePicture(imageId, () => imageId switch
    {
        SymbolPalette.UpImage => SymbolPictures.Arrow(SymbolPictures.ArrowKind.Up, SymbolPalette.SlotSize),
        SymbolPalette.DownImage => SymbolPictures.Arrow(SymbolPictures.ArrowKind.Down, SymbolPalette.SlotSize),
        SymbolPalette.MoreImage => SymbolPictures.Arrow(SymbolPictures.ArrowKind.More, SymbolPalette.SlotSize),
        _ => SymbolPalette.FromImageId(imageId) switch
        {
            ({ } c, true) => SymbolPictures.Icon(c, SymbolPalette.GroupIconSize),
            ({ } c, false) => SymbolPictures.Cell(c, SymbolPalette.ItemSize),
            _ => null,
        },
    });

    private static object? SafePicture(string name, Func<object?> draw)
    {
        try { return draw(); }
        catch (Exception ex)
        {
            Log.Error($"Drawing the ribbon picture {name} failed", ex);
            return null;
        }
    }

    /// <summary>The symbol a grid cell shows at the current scroll position, if any.</summary>
    private Symbol? SlotSymbol(IRibbonControl control) =>
        SymbolPalette.FromSlotId(control.Id) is var (row, column) ? SymbolPalette.At(_symbolTopRow + row, column) : null;

    public object? GetSlotImage(IRibbonControl control) => SafePicture(control.Id, () =>
        SlotSymbol(control) is { } s ? SymbolPictures.Cell(s.Character, SymbolPalette.SlotSize) : SymbolPictures.Blank(SymbolPalette.SlotSize));

    public string GetSlotLabel(IRibbonControl control) => SlotSymbol(control)?.Character.ToString() ?? " ";

    public string GetSlotScreentip(IRibbonControl control) => SlotSymbol(control) is { } s ? SymbolPalette.ScreenTip(s) : "";

    public string GetSlotSupertip(IRibbonControl control) => SlotSymbol(control) is { } s ? SymbolPalette.SuperTip(s) : "";

    public bool GetSlotEnabled(IRibbonControl control) => SlotSymbol(control) is not null;

    /// <summary>A grid cell clicked: its symbol is typed at the cursor.</summary>
    public void OnSlot(IRibbonControl control)
    {
        if (SlotSymbol(control) is { } s)
            Run("Symbol", () => Controller.InsertSymbol(s.Character.ToString()));
    }

    /// <summary>A symbol picked from the drop-down with all of them.</summary>
    public void OnSymbol(IRibbonControl control, string selectedId, int selectedIndex)
    {
        if (SymbolPalette.FromItemId(selectedId) is { } c)
            Run("Symbol", () => Controller.InsertSymbol(c.ToString()));
    }

    public bool GetSymbolsUpEnabled(IRibbonControl control) => _symbolTopRow > 0;

    public bool GetSymbolsDownEnabled(IRibbonControl control) => _symbolTopRow < SymbolPalette.MaxTopRow;

    public void OnSymbolsUp(IRibbonControl control) => ScrollSymbols(-1);

    public void OnSymbolsDown(IRibbonControl control) => ScrollSymbols(+1);

    private void ScrollSymbols(int rows)
    {
        _symbolTopRow = Math.Clamp(_symbolTopRow + rows, 0, SymbolPalette.MaxTopRow);
        if (_ribbon is null)
            return;
        foreach (var id in SymbolPalette.SlotIds())
            _ribbon.InvalidateControl(id);
        _ribbon.InvalidateControl("CalcpadSymbolsUp");
        _ribbon.InvalidateControl("CalcpadSymbolsDown");
    }

    /// <summary>"Edit Calcpad Source" on results and "Calculate Calcpad" on source text, in each right-click menu.</summary>
    private static string ContextMenusXml()
    {
        var sb = new StringBuilder("<contextMenus>");
        foreach (var menu in ContextMenus)
        {
            sb.Append($"<contextMenu idMso=\"{menu}\">")
              .Append($"<button id=\"CalcpadEditSource_{menu}\" tag=\"edit\" label=\"Edit Calcpad Source\" imageMso=\"EquationLinearFormat\" ")
              .Append("onAction=\"OnEditSource\" getVisible=\"GetContextItemVisible\" />")
              .Append($"<button id=\"CalcpadCalculate_{menu}\" tag=\"calculate\" label=\"Calculate Calcpad\" imageMso=\"CalculateNow\" ")
              .Append("onAction=\"OnRecalculate\" getVisible=\"GetContextItemVisible\" />")
              .Append($"<button id=\"CalcpadBlockToText_{menu}\" tag=\"block\" label=\"Convert Calcpad Block to Word Text\" imageMso=\"ContentControlText\" ")
              .Append("onAction=\"OnConvertBlock\" getVisible=\"GetContextItemVisible\" />")
              .Append("</contextMenu>");
        }
        return sb.Append("</contextMenus>").ToString();
    }

    // ---- Ribbon callbacks (called by name through IDispatch) ----------------------------

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbon = ribbon;

    public void OnEvaluate(IRibbonControl control) => Run("Convert", () => Controller.Evaluate());

    public void OnRecalculate(IRibbonControl control) => Run("Calculate", () => Controller.Recalculate());

    public void OnEditSource(IRibbonControl control) => Run("Edit Source", () => Controller.ToggleEditSource());

    public void OnLock(IRibbonControl control) => Run("Lock", () => Controller.SetLocked(true));

    public void OnUnlock(IRibbonControl control) => Run("Unlock", () => Controller.SetLocked(false));

    public void OnCancel(IRibbonControl control) => Run("Cancel", () => Controller.Cancel());

    public bool GetCancelEnabled(IRibbonControl control) { Log.Trace("ribbon GetCancelEnabled"); return _controller?.CanCancel == true; }

    public void OnMakeBlock(IRibbonControl control) => Run("Calcpad Block", () => Controller.MakeBlock());

    public void OnNextError(IRibbonControl control) => Run("Next Error", () => Controller.NextError());

    public void OnGoToError(IRibbonControl control) =>
        Run("Go to error", () => Controller.GoToError(int.TryParse(control.Tag, out var i) ? i : 0));

    public string GetErrorListLabel(IRibbonControl control) =>
        SafeErrors().Count is var n and > 0 ? $"Errors ({n})" : "No Errors";

    public bool GetHasErrors(IRibbonControl control) => SafeErrors().Count > 0;

    public string GetErrorListContent(IRibbonControl control)
    {
        var sb = new StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
        var errors = SafeErrors();
        for (var i = 0; i < errors.Count && i < MaxErrorMenuItems; ++i)
        {
            var label = Clip(errors[i].Source.Length == 0 ? "(empty)" : errors[i].Source, 60);
            sb.Append($"<button id=\"CalcpadError{i}\" tag=\"{i}\" onAction=\"OnGoToError\" ")
              .Append($"label=\"{SecurityElement.Escape(label)}\" ")
              .Append($"supertip=\"{SecurityElement.Escape(Clip(errors[i].Message, 500))}\" />");
        }
        if (errors.Count > MaxErrorMenuItems)
            sb.Append($"<button id=\"CalcpadErrorMore\" enabled=\"false\" label=\"…and {errors.Count - MaxErrorMenuItems} more (use Next Error)\" />");
        return sb.Append("</menu>").ToString();
    }

    /// <summary>Right-click menu: Edit Source on a result, Calculate on source text, nothing elsewhere.</summary>
    public bool GetContextItemVisible(IRibbonControl control)
    {
        if (control.Tag == "block")
            return SafeBlockAtSelection();

        var kind = SafeKindAtSelection();
        return control.Tag switch
        {
            "edit" => kind == CalcpadKind.Result,
            "calculate" => kind == CalcpadKind.Source,
            _ => false,
        };
    }

    public void OnConvertBlock(IRibbonControl control) => Run("Block to Text", () => Controller.ConvertBlockToText());

    // Math font: the starters, marked "(not installed)" where needed, plus the document's own font if it's another.

    public int GetMathFontCount(IRibbonControl control) => SafeFontChoices().Count;

    public string GetMathFontLabel(IRibbonControl control, int index) =>
        SafeFontChoices() is var choices && index < choices.Count
            ? choices[index].Font.Label + (choices[index].Installed ? "" : " (not installed)")
            : "";

    public string GetMathFontId(IRibbonControl control, int index) => $"CalcpadMathFont{index}";

    public string GetMathFontSupertip(IRibbonControl control, int index) =>
        SafeFontChoices() is var choices && index < choices.Count ? choices[index].Font.Note : "";

    public int GetMathFontIndex(IRibbonControl control)
    {
        var current = SafeCurrentFont();
        var index = SafeFontChoices().ToList().FindIndex(c => c.Font.Family.Equals(current, StringComparison.OrdinalIgnoreCase));
        return Math.Max(0, index);
    }

    public void OnMathFontChanged(IRibbonControl control, string selectedId, int selectedIndex)
    {
        var choices = SafeFontChoices();
        if (selectedIndex < 0 || selectedIndex >= choices.Count)
            return;

        var family = choices[selectedIndex].Font.Family;
        Run("Math font", () =>
        {
            try
            {
                return Controller.SetMathFont(family, allowDownload: false);
            }
            catch (FontNotInstalledException ex)
            {
                if (!Confirm($"{ex.Font.Family} is not installed on this computer.\n\n{ex.Font.Note}\n\n" +
                             $"Download it and install it for you (no administrator rights needed)?\nFrom: {new Uri(ex.Font.DownloadUrl!).Host}"))
                    return string.Empty;
                return Controller.SetMathFont(family, allowDownload: true);
            }
        });
    }

    public int GetDecimalsIndex(IRibbonControl control) => SafeSettings().Decimals;

    public void OnDecimalsChanged(IRibbonControl control, string selectedId, int selectedIndex) =>
        Run("Decimals", () => Controller.ChangeSettings(s => s with { Decimals = selectedIndex }));

    public int GetAnglesIndex(IRibbonControl control) => (int)SafeSettings().Angles;

    public void OnAnglesChanged(IRibbonControl control, string selectedId, int selectedIndex) =>
        Run("Angles", () => Controller.ChangeSettings(s => s with { Angles = (AngleUnits)selectedIndex }));

    public bool GetAutomaticPressed(IRibbonControl control) => SafeSettings().Automatic;

    /// <summary>"Calculate", or a reminder when the results are out of date (automatic calculation off).</summary>
    public string GetRecalculateLabel(IRibbonControl control)
    {
        try { return _controller?.IsOutOfDate() == true ? "Calculate (out of date)" : "Calculate"; }
        catch (COMException) { return "Calculate"; }
    }

    public void OnAutomaticChanged(IRibbonControl control, bool pressed) =>
        Run("Calculate automatically", () => Controller.ChangeSettings(s => s with { Automatic = pressed }));

    public bool GetSubstitutePressed(IRibbonControl control) => SafeSettings().Substitute;

    public void OnSubstituteChanged(IRibbonControl control, bool pressed) =>
        Run("Substitute values", () => Controller.ChangeSettings(s => s with { Substitute = pressed }));

    public bool GetComplexPressed(IRibbonControl control) => SafeSettings().Complex;

    public void OnComplexChanged(IRibbonControl control, bool pressed) =>
        Run("Complex numbers", () => Controller.ChangeSettings(s => s with { Complex = pressed }));

    public bool GetPlainStylePressed(IRibbonControl control) => SafeSettings().Style == RenderStyle.PlainMath;

    public void OnPlainStyleChanged(IRibbonControl control, bool pressed) =>
        Run("Plain Word math", () => Controller.ChangeSettings(s => s with { Style = pressed ? RenderStyle.PlainMath : RenderStyle.Calcpad }));

    // ---- Automation API (macros and tests) ----------------------------------------------
    //
    // Calls from other processes arrive on COM worker threads; each is run on Word's UI thread,
    // so it never overlaps a command, an event or background results being drawn.

    public string Version => typeof(Connect).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "?";

    public string Evaluate() => Api(() => Controller.Evaluate());

    public string EvaluateOrEditSource() => Api(() => Controller.EvaluateOrEditSource());

    public string Recalculate() => Api(() => Controller.Recalculate());

    public string EditSource() => Api(() => Controller.ToggleEditSource());

    /// <summary>Lock and Unlock: the Calcpad expressions and blocks in the selection.</summary>
    public string Lock() => Api(() => Controller.SetLocked(true));

    public string Unlock() => Api(() => Controller.SetLocked(false));

    /// <summary>What the Enter key does (the keyboard hook can't be driven by automation).</summary>
    public string PressEnter() => Api(PressEnterCore);

    /// <summary>True when Enter at the cursor would be taken by Calcpad.</summary>
    public bool EnterLeavesControl => _enterLeavesControl;

    /// <summary>A calculation or font installation is still running in the background.</summary>
    public bool IsWorking => Controller.IsWorking; // not via the UI thread: polling it would keep interrupting Word

    /// <summary>The status of the last finished recalculation (also after background work).</summary>
    public string LastStatus => Controller.LastStatus;

    /// <summary>What the Cancel button (and Esc) does.</summary>
    public string Cancel() => Api(() => Controller.Cancel());

    /// <summary>Whether the Cancel button is enabled.</summary>
    public bool CanCancel => Controller.CanCancel;

    /// <summary>The math fonts on offer, "|"-separated, each marked "+" if installed or "-" if not.</summary>
    public string MathFontList => UiThread.Invoke(() => string.Join("|", Controller.MathFontChoices().Select(c => (c.Installed ? "+" : "-") + c.Font.Family)));

    public string MathFont => UiThread.Invoke(() => Controller.CurrentMathFont());

    /// <summary>Sets the math font; a free font that isn't installed is downloaded and installed first (in the background).</summary>
    public string SetMathFont(string family) => Api(() => Controller.SetMathFont(family, allowDownload: true));

    public string MakeBlock() => Api(() => Controller.MakeBlock());

    /// <summary>What picking a symbol from the Calcpad tab's galleries does.</summary>
    public string InsertSymbol(string symbol) => Api(() => Controller.InsertSymbol(symbol));

    /// <summary>
    /// The picture LoadImage gives the ribbon for an image name, as "bitmap W x H" (pixels at
    /// 96 dpi), or "none". Described here because the picture itself doesn't cross processes.
    /// </summary>
    public string DescribeImage(string imageId) => UiThread.Invoke(() =>
    {
        if (LoadImage(imageId) is not { } picture)
            return "none";

        int Get(string name) => Convert.ToInt32(picture.GetType().InvokeMember(name, BindingFlags.GetProperty, null, picture, null));
        static int Pixels(int himetric) => (int)Math.Round(himetric * 96 / 2540.0);
        return $"{(Get("Type") == 1 ? "bitmap" : $"type {Get("Type")}")} {Pixels(Get("Width"))} x {Pixels(Get("Height"))}";
    });

    public string ConvertBlockToText() => Api(() => Controller.ConvertBlockToText());

    public string NextError() => Api(() => Controller.NextError());

    public int ErrorCount => UiThread.Invoke(() => Controller.Errors().Count);

    /// <summary>"result", "source" or "none": what the right-click menu would offer at the cursor.</summary>
    public string CalcpadAtSelection => UiThread.Invoke(() => SafeKindAtSelection().ToString().ToLowerInvariant());

    public string GetSettings() => UiThread.Invoke(() => Controller.GetSettings().Serialize());

    public string SetSettings(string settings) => Api(() => Controller.ChangeSettings(_ => DocumentSettings.Parse(settings)));

    /// <summary>Automation commands: run on the UI thread, errors go back to the caller, the cursor state is refreshed like after a command.</summary>
    private string Api(Func<string> action) => UiThread.Invoke(() =>
    {
        try { return action(); }
        catch (Exception ex) when (ex is not UserMessageException)
        {
            Log.Error("Automation call failed", ex);
            throw;
        }
        finally { RefreshCursorState(); }
    });

    // ---- Events -------------------------------------------------------------------------

    /// <summary>Leaving edited Calcpad source evaluates it, right after Word finishes moving the cursor.</summary>
    private void OnControlExited(Word.Document doc, Word.ContentControl cc)
    {
        if (_controller?.NeedsEvaluationAfterExit(cc) != true)
            return;

        UiThread.Post(() =>
        {
            if (IsOpen(doc))
                Run("Calculate", () => Controller.AfterExit(doc, cc));
        }, key: "evaluate-after-exit");
    }

    /// <summary>Double-clicking a result opens its source.</summary>
    private bool OnDoubleClick(Word.Selection selection)
    {
        if (_controller?.IsCalcpadResult(selection.Range) != true)
            return false;

        UiThread.Post(() => Run("Edit Source", () => Controller.ToggleEditSource()), key: "edit-source");
        return true; // cancel Word's own double-click (word selection)
    }

    /// <summary>
    /// The Enter key, taken from Word while the cursor was in an expression. Re-checked here; if
    /// the cursor has moved somewhere ordinary meanwhile, it types a normal Enter instead.
    /// </summary>
    private string PressEnterCore()
    {
        if (Controller.LeaveWithEnter())
            return "left";

        _app?.Selection.TypeParagraph();
        return "typed";
    }

    /// <summary>Word refuses most calls during some events, such as a document closing.</summary>
    private const int NotAvailableInThisEvent = unchecked((int)0x800A1835);

    private void RefreshCursorState() => RefreshCursorState(retried: false);

    private void RefreshCursorState(bool retried)
    {
        try
        {
            _enterLeavesControl = _app is not null && _controller is not null && _app.Documents.Count > 0 &&
                                  _controller.EnterLeavesControlAt(_app.Selection.Range);
        }
        catch (COMException ex) when (ex.ErrorCode == NotAvailableInThisEvent && !retried)
        {
            _enterLeavesControl = false; // never take Enter when unsure; look again once the event is over
            UiThread.Post(() => RefreshCursorState(retried: true), key: "refresh-cursor-state");
        }
        catch (Exception ex)
        {
            _enterLeavesControl = false;
            Log.Error("Checking the cursor position failed", ex);
        }
    }

    private void OnDocumentChange()
    {
        RefreshCursorState();
        _ribbon?.Invalidate();
    }

    /// <summary>Results calculated in the background (or a font download) are in.</summary>
    private void OnBackgroundFinished(string status)
    {
        try
        {
            if (_app is not null && status.Length > 0)
                _app.StatusBar = status;
        }
        catch (COMException) { }
        RefreshCursorState();
        _ribbon?.Invalidate();
        if (status.Contains("could not be installed", StringComparison.Ordinal))
            ShowMessage(status);
    }

    private IReadOnlyList<(MathFontInfo Font, bool Installed)> SafeFontChoices()
    {
        Log.Trace("ribbon SafeFontChoices");
        try { return _controller?.MathFontChoices() ?? []; }
        catch (Exception ex) { Log.Error("Listing math fonts failed", ex); return []; }
    }

    private string SafeCurrentFont()
    {
        Log.Trace("ribbon SafeCurrentFont");
        try { return _controller?.CurrentMathFont() ?? MathFonts.Default; }
        catch (Exception ex) { Log.Error("Reading the math font failed", ex); return MathFonts.Default; }
    }

    /// <summary>Word is about to show a context menu: make it re-check which Calcpad items apply.</summary>
    private void OnRightClick()
    {
        if (_ribbon is null)
            return;

        foreach (var id in ContextItemIds)
            _ribbon.InvalidateControl(id);
    }

    // ---- Plumbing -----------------------------------------------------------------------

    /// <summary>Runs a user command: status bar on success, a message for problems, ribbon refreshed after.</summary>
    private void Run(string command, Func<string> action)
    {
        try
        {
            var status = action();
            if (_app is not null && status.Length > 0)
                _app.StatusBar = status;
        }
        catch (UserMessageException ex)
        {
            ShowMessage(ex.Message);
        }
        catch (Exception ex)
        {
            Log.Error($"{command} failed", ex);
            ShowMessage($"Calcpad {command} failed: {ex.Message}\n\nDetails: {Log.FilePath}");
        }
        finally
        {
            RefreshCursorState();
            _ribbon?.Invalidate();
        }
    }

    private DocumentSettings SafeSettings()
    {
        Log.Trace("ribbon SafeSettings");
        try { return _controller?.GetSettings() ?? new DocumentSettings(); }
        catch (Exception ex) { Log.Error("Reading settings failed", ex); return new DocumentSettings(); }
    }

    private CalcpadKind SafeKindAtSelection()
    {
        Log.Trace("ribbon SafeKindAtSelection");
        try { return _app is null || _controller is null || _app.Documents.Count == 0 ? CalcpadKind.None : _controller.KindAt(_app.Selection.Range); }
        catch (Exception ex) { Log.Error("Checking the cursor position failed", ex); return CalcpadKind.None; }
    }

    private bool SafeBlockAtSelection()
    {
        Log.Trace("ribbon SafeBlockAtSelection");
        try { return _app is not null && _controller is not null && _app.Documents.Count > 0 && _controller.IsBlockAt(_app.Selection.Range); }
        catch (Exception ex) { Log.Error("Checking the cursor position failed", ex); return false; }
    }

    private IReadOnlyList<CalcpadError> SafeErrors()
    {
        Log.Trace("ribbon SafeErrors");
        try { return _controller?.Errors() ?? []; }
        catch (Exception ex) { Log.Error("Listing errors failed", ex); return []; }
    }

    private static bool IsOpen(Word.Document doc)
    {
        try { _ = doc.Name; return true; }
        catch (COMException) { return false; }
    }

    private static string Clip(string s, int max) => s.Length <= max ? s : s[..(max - 1)] + "…";

    private void ShowMessage(string text)
    {
        // A Word nobody can see (automation) would wait forever on a message nobody can close.
        if (!IsWordVisible())
        {
            Log.Info($"Message not shown (Word is hidden): {text}");
            return;
        }
        _ = MessageBoxW(OwnerWindow(), text, "Calcpad", 0x40 /* MB_ICONINFORMATION */);
    }

    private bool IsWordVisible()
    {
        try { return _app?.Visible == true; }
        catch (COMException) { return false; }
    }

    private bool Confirm(string question) =>
        MessageBoxW(OwnerWindow(), question, "Calcpad", 0x4 | 0x20 /* MB_YESNO | MB_ICONQUESTION */) == 6 /* IDYES */;

    private IntPtr OwnerWindow()
    {
        try { return new IntPtr(_app?.ActiveWindow.Hwnd ?? 0); }
        catch (COMException) { return IntPtr.Zero; }
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);
}
