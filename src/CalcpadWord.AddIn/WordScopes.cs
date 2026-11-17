using Word = Microsoft.Office.Interop.Word;

namespace CalcpadWord.AddIn;

/// <summary>Groups every change made during one command into a single Undo step.</summary>
internal sealed class UndoScope : IDisposable
{
    private readonly Word.UndoRecord? _record;

    public UndoScope(Word.Application app, string name)
    {
        try
        {
            var record = app.UndoRecord;
            if (!record.IsRecordingCustomRecord)
            {
                record.StartCustomRecord(name);
                _record = record;
            }
        }
        catch (Exception ex)
        {
            Log.Error("Could not start an undo record", ex);
        }
    }

    public void Dispose()
    {
        try { _record?.EndCustomRecord(); }
        catch (Exception ex) { Log.Error("Could not end the undo record", ex); }
    }
}

/// <summary>Turns screen updating off while many controls are redrawn.</summary>
internal sealed class ScreenUpdatingScope : IDisposable
{
    private readonly Word.Application _app;
    private readonly bool _previous;

    public ScreenUpdatingScope(Word.Application app)
    {
        _app = app;
        _previous = app.ScreenUpdating;
        app.ScreenUpdating = false;
    }

    public void Dispose()
    {
        try { _app.ScreenUpdating = _previous; }
        catch (Exception ex) { Log.Error("Could not restore screen updating", ex); }
    }
}
