using Word = Microsoft.Office.Interop.Word;

namespace CalcpadWord.AddIn;

/// <summary>
/// Word events the add-in reacts to: leaving a content control (per document), double- and
/// right-clicking, the selection moving, and switching documents.
/// Every subscription is best-effort: a failure is logged and the add-in keeps working without it.
/// </summary>
internal sealed class WordEvents : IDisposable
{
    private readonly Word.Application _app;
    private readonly Action<Word.Document, Word.ContentControl> _controlExited;
    private readonly Func<Word.Selection, bool> _doubleClicked;
    private readonly Action _rightClicked;
    private readonly Action _selectionChanged;
    private readonly Action _documentChanged;
    private readonly HashSet<Word.Document> _watched = new(ReferenceEqualityComparer.Instance);
    private readonly List<Action> _unsubscribe = [];

    public WordEvents(
        Word.Application app,
        Action<Word.Document, Word.ContentControl> controlExited,
        Func<Word.Selection, bool> doubleClicked,
        Action rightClicked,
        Action selectionChanged,
        Action documentChanged)
    {
        _app = app;
        _controlExited = controlExited;
        _doubleClicked = doubleClicked;
        _rightClicked = rightClicked;
        _selectionChanged = selectionChanged;
        _documentChanged = documentChanged;
    }

    public void Start()
    {
        var events = (Word.ApplicationEvents4_Event)_app;
        Try("DocumentChange", () =>
        {
            Word.ApplicationEvents4_DocumentChangeEventHandler h = OnDocumentChange;
            events.DocumentChange += h;
            _unsubscribe.Add(() => events.DocumentChange -= h);
        });
        Try("DocumentOpen", () =>
        {
            Word.ApplicationEvents4_DocumentOpenEventHandler h = Watch;
            events.DocumentOpen += h;
            _unsubscribe.Add(() => events.DocumentOpen -= h);
        });
        Try("NewDocument", () =>
        {
            Word.ApplicationEvents4_NewDocumentEventHandler h = Watch;
            events.NewDocument += h;
            _unsubscribe.Add(() => events.NewDocument -= h);
        });
        Try("WindowBeforeDoubleClick", () =>
        {
            Word.ApplicationEvents4_WindowBeforeDoubleClickEventHandler h = OnDoubleClick;
            events.WindowBeforeDoubleClick += h;
            _unsubscribe.Add(() => events.WindowBeforeDoubleClick -= h);
        });
        Try("WindowBeforeRightClick", () =>
        {
            Word.ApplicationEvents4_WindowBeforeRightClickEventHandler h = OnRightClick;
            events.WindowBeforeRightClick += h;
            _unsubscribe.Add(() => events.WindowBeforeRightClick -= h);
        });
        Try("WindowSelectionChange", () =>
        {
            Word.ApplicationEvents4_WindowSelectionChangeEventHandler h = _ =>
            {
                Log.Trace("SelectionChange start");
                try { _selectionChanged(); }
                catch (Exception ex) { Log.Error("Handling a selection change failed", ex); }
                Log.Trace("SelectionChange end");
            };
            events.WindowSelectionChange += h;
            _unsubscribe.Add(() => events.WindowSelectionChange -= h);
        });

        // Documents that were already open when the add-in connected.
        Try("open documents", () =>
        {
            var docs = _app.Documents;
            for (int i = 1, n = docs.Count; i <= n; ++i)
                Watch(docs[i]);
        });
    }

    public void Dispose()
    {
        foreach (var unsubscribe in _unsubscribe)
            Try("unsubscribe", unsubscribe);
        _unsubscribe.Clear();
        _watched.Clear();
    }

    private void OnDocumentChange()
    {
        Log.Trace("DocumentChange start");
        // Also catches documents opened in ways that raise no DocumentOpen (e.g. from Explorer at startup).
        Try("watch active document", () =>
        {
            if (_app.Documents.Count > 0)
                Watch(_app.ActiveDocument);
        });
        _documentChanged();
        Log.Trace("DocumentChange end");
    }

    private void Watch(Word.Document doc)
    {
        // The add-in's own scratch document lives for a moment: nothing to watch.
        if (DocumentController.MakingScratch || !_watched.Add(doc))
            return;

        Try("ContentControlOnExit", () =>
        {
            var events = (Word.DocumentEvents2_Event)doc;
            Word.DocumentEvents2_ContentControlOnExitEventHandler h = (Word.ContentControl cc, ref bool cancel) =>
            {
                try { _controlExited(doc, cc); }
                catch (Exception ex) { Log.Error("Handling ContentControlOnExit failed", ex); }
            };
            events.ContentControlOnExit += h;
            _unsubscribe.Add(() => events.ContentControlOnExit -= h);
        });
    }

    private void OnDoubleClick(Word.Selection selection, ref bool cancel)
    {
        try
        {
            if (_doubleClicked(selection))
                cancel = true;
        }
        catch (Exception ex)
        {
            Log.Error("Handling double-click failed", ex);
        }
    }

    // Runs before Word shows its context menu: lets the ribbon re-check which Calcpad items to show.
    private void OnRightClick(Word.Selection selection, ref bool cancel)
    {
        try { _rightClicked(); }
        catch (Exception ex) { Log.Error("Handling right-click failed", ex); }
    }

    private static void Try(string what, Action action)
    {
        try { action(); }
        catch (Exception ex) { Log.Error($"Word event setup failed: {what}", ex); }
    }
}
