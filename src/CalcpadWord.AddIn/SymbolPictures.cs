using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace CalcpadWord.AddIn;

/// <summary>
/// Pictures for the Symbols group: framed cells as on Word's equation Symbols gallery, the
/// group's icon, and the ▲ ▼ ▾ buttons. Symbols are drawn with GDI (no extra libraries) in
/// Cambria Math, as Word draws its equation symbols, eight times larger and then scaled down:
/// GDI's hinting at small sizes would make them bold and jagged, while this gives thin, smooth
/// strokes like Word's. Grey ink and frames as Word's, an alpha channel, at the screen's real
/// scale. Office takes them as IPictureDisp. Each is drawn once.
/// </summary>
internal static class SymbolPictures
{
    private const string MathFont = "Cambria Math";
    private const string FallbackFont = "Segoe UI Symbol"; // for ‱, which Cambria Math lacks
    private const int Oversampling = 8;

    // Word's own gallery, measured at 125 %: glyphs no darker than #5E, frames #808080.
    private const byte LightInk = 0x55, LightFrame = 0x80;
    private const byte DarkInk = 0xE6, DarkFrame = 0x76;

    public enum ArrowKind { Up, Down, More }

    // Ribbon callbacks and LoadImage all run on Word's UI thread.
    private static readonly Dictionary<string, object> Cache = new(StringComparer.Ordinal);
    private static double? _scale;

    /// <summary>A cell: the symbol on a common baseline (° stays high, β descends), centred across, in a thin frame.</summary>
    public static object Cell(char c, int logicalSize) =>
        Cached($"cell {c} {logicalSize}", logicalSize, size => (Draw(c, size, icon: false), Frame(size)));

    /// <summary>The group's icon, shown when the ribbon is too narrow and folds the group: the symbol alone.</summary>
    public static object Icon(char c, int logicalSize) =>
        Cached($"icon {c} {logicalSize}", logicalSize, size => (Draw(c, size, icon: true), null));

    /// <summary>An empty cell (the end of the last row).</summary>
    public static object Blank(int logicalSize) =>
        Cached($"blank {logicalSize}", logicalSize, size => (new byte[size * size], null));

    /// <summary>▲, ▼, or ▼ under a bar (all of them), as on Word's gallery scroll bar.</summary>
    public static object Arrow(ArrowKind kind, int logicalSize) =>
        Cached($"arrow {kind} {logicalSize}", logicalSize, size => (DrawArrow(kind, size), null));

    private static object Cached(string key, int logicalSize, Func<int, (byte[] Ink, byte[]? Frame)> draw)
    {
        if (!Cache.TryGetValue(key, out var picture))
        {
            var size = Math.Max(8, (int)Math.Round(logicalSize * DpiScale()));
            var (ink, frame) = draw(size);
            var dark = IsDarkTheme();
            picture = ToPicture(size, ink, dark ? DarkInk : LightInk, frame, dark ? DarkFrame : LightFrame);
            Cache[key] = picture;
        }
        return picture;
    }

    private static byte[] Frame(int size)
    {
        var frame = new byte[size * size];
        for (var i = 0; i < size; ++i)
            frame[i] = frame[(size - 1) * size + i] = frame[i * size] = frame[i * size + size - 1] = 255;
        return frame;
    }

    /// <summary>
    /// The glyph's coverage (0–255) in a size × size square, rows from the top. Capitals are
    /// centred on the baseline; a glyph too wide is drawn smaller, and a very short one (°, ′,
    /// ‾) up to 1.5 times larger, by its height, so ′ ″ ‴ ⁗ stay alike.
    /// </summary>
    private static byte[] Draw(char c, int size, bool icon)
    {
        const int k = Oversampling;
        var result = new byte[size * size];
        var margin = (icon ? 0 : 1) + Math.Max(1, size / 16); // clear of the frame
        // As Word's cells (24 px at 100 %); the ribbon's smaller ones (16 px) get relatively larger glyphs.
        var em = size * (icon ? 0.9 : size < 22 * DpiScale() ? 0.72 : 0.62) * k;
        var glyph = Render(c, size * k, em);
        if (glyph.Right < 0)
            return result; // nothing drawn

        var shortest = (icon ? 0.45 : 0.34) * size * k;
        var grow = glyph.Height < shortest && (icon || !char.IsLetter(c)) // letters stay alike (α beside β)
            ? Math.Min(icon ? 3.0 : 1.5, shortest / glyph.Height)
            : 1.0;
        var scale = Math.Min(grow, (size - 2.0 * margin) * k / glyph.Width); // never wider than the cell
        if (Math.Abs(scale - 1) > 0.02)
            glyph = Render(c, size * k, em * scale);

        // Placed in oversampled pixels: centred across, the baseline on a whole pixel.
        var dx = (size * k - glyph.Width) / 2 - glyph.Left;
        var dy = (int)Math.Round(size * (icon ? 0.75 : 0.70)) * k - glyph.Baseline;
        if (glyph.Bottom + dy >= (size - margin) * k)
            dy = (size - margin) * k - 1 - glyph.Bottom;
        if (glyph.Top + dy < margin * k)
            dy = margin * k - glyph.Top;

        var sums = new int[size * size];
        for (var y = glyph.Top; y <= glyph.Bottom; ++y)
        {
            var sy = y + dy;
            if (sy < 0 || sy >= size * k)
                continue;
            for (var x = glyph.Left; x <= glyph.Right; ++x)
            {
                var sx = x + dx;
                if (sx >= 0 && sx < size * k)
                    sums[sy / k * size + sx / k] += glyph.Coverage(x, y);
            }
        }
        for (var i = 0; i < sums.Length; ++i)
            result[i] = (byte)Math.Min(255, sums[i] / (k * k));
        return result;
    }

    /// <summary>A small triangle pointing up or down (under a bar for More), antialiased with 4 × 4 samples a pixel.</summary>
    private static byte[] DrawArrow(ArrowKind kind, int size)
    {
        double width = size * 0.46, height = size * 0.24, cx = size / 2.0;
        var cy = size / 2.0 + (kind == ArrowKind.More ? size * 0.09 : 0);
        double top = cy - height / 2, bottom = cy + height / 2;
        var barTop = top - size * 0.19;
        var barBottom = barTop + Math.Max(1.0, size / 14.0);

        bool Inside(double x, double y)
        {
            if (kind == ArrowKind.More && y >= barTop && y < barBottom && Math.Abs(x - cx) <= width / 2)
                return true;
            if (y < top || y > bottom)
                return false;
            var fromApex = kind == ArrowKind.Up ? y - top : bottom - y; // 0 at the point
            return Math.Abs(x - cx) <= width / 2 * fromApex / height;
        }

        var result = new byte[size * size];
        for (var y = 0; y < size; ++y)
            for (var x = 0; x < size; ++x)
            {
                var hits = 0;
                for (var sy = 0; sy < 4; ++sy)
                    for (var sx = 0; sx < 4; ++sx)
                        if (Inside(x + (sx + 0.5) / 4, y + (sy + 0.5) / 4))
                            ++hits;
                result[y * size + x] = (byte)(hits * 255 / 16);
            }
        return result;
    }

    /// <summary>A glyph drawn in a scratch bitmap, with the box its ink covers (Right is -1 if none).</summary>
    private sealed record Glyph(byte[] Pixels, int Stride, int Baseline, int Left, int Top, int Right, int Bottom)
    {
        public int Width => Right - Left + 1;
        public int Height => Bottom - Top + 1;

        public byte Coverage(int x, int y)
        {
            var i = (y * Stride + x) * 4;
            return Math.Max(Pixels[i], Math.Max(Pixels[i + 1], Pixels[i + 2]));
        }
    }

    /// <summary>
    /// Draws the glyph white on black in a scratch bitmap two cells wide and high, so that
    /// antialiasing makes the brightness the coverage.
    /// </summary>
    private static Glyph Render(char c, int size, double emPixels)
    {
        int width = 2 * size, height = 2 * size, baseline = (int)(1.3 * size);
        var em = Math.Max(4, (int)Math.Round(emPixels));
        var header = new BITMAPINFOHEADER { biSize = 40, biWidth = width, biHeight = -height, biPlanes = 1, biBitCount = 32 };
        var dc = CreateCompatibleDC(IntPtr.Zero);
        var bitmap = CreateDIBSection(dc, ref header, 0, out var bits, IntPtr.Zero, 0);
        var font = IntPtr.Zero;
        var scratch = new byte[width * height * 4];
        try
        {
            if (bitmap == IntPtr.Zero)
                throw new InvalidOperationException("Could not create a bitmap for a ribbon symbol.");
            SelectObject(dc, bitmap);
            font = CreateFont(em, MathFont);
            SelectObject(dc, font);
            if (!HasGlyph(dc, c))
            {
                var fallback = CreateFont(em, FallbackFont);
                SelectObject(dc, fallback);
                DeleteObject(font);
                font = fallback;
            }
            SetTextColor(dc, 0xFFFFFF);
            SetBkMode(dc, Transparent);
            SetTextAlign(dc, TextAlignBaseline);
            TextOutW(dc, size / 4, baseline, c.ToString(), 1);
            GdiFlush();
            Marshal.Copy(bits, scratch, 0, scratch.Length);
        }
        finally
        {
            DeleteDC(dc); // before the objects it held
            if (font != IntPtr.Zero)
                DeleteObject(font);
            if (bitmap != IntPtr.Zero)
                DeleteObject(bitmap);
        }

        int left = width, right = -1, top = height, bottom = -1;
        for (var y = 0; y < height; ++y)
            for (var x = 0; x < width; ++x)
            {
                var i = (y * width + x) * 4;
                if (scratch[i] != 0 || scratch[i + 1] != 0 || scratch[i + 2] != 0)
                {
                    left = Math.Min(left, x);
                    right = Math.Max(right, x);
                    top = Math.Min(top, y);
                    bottom = Math.Max(bottom, y);
                }
            }
        return new Glyph(scratch, width, baseline, left, top, right, bottom);
    }

    /// <summary>
    /// A 32-bit bitmap with premultiplied alpha (the ink over the frame, each a coverage and a
    /// grey), wrapped as an OLE picture that owns it.
    /// </summary>
    private static object ToPicture(int size, byte[] ink, byte inkGrey, byte[]? frame, byte frameGrey)
    {
        var header = new BITMAPINFOHEADER { biSize = 40, biWidth = size, biHeight = size, biPlanes = 1, biBitCount = 32 }; // bottom-up
        var bitmap = CreateDIBSection(IntPtr.Zero, ref header, 0, out var bits, IntPtr.Zero, 0);
        if (bitmap == IntPtr.Zero)
            throw new InvalidOperationException("Could not create a bitmap for a ribbon symbol.");

        var pixels = new byte[size * size * 4];
        for (var y = 0; y < size; ++y)
            for (var x = 0; x < size; ++x)
            {
                var a = ink[y * size + x] / 255.0;
                var f = (frame?[y * size + x] ?? 0) / 255.0 * (1 - a);
                var i = ((size - 1 - y) * size + x) * 4;
                var value = (byte)Math.Round(inkGrey * a + frameGrey * f);
                pixels[i] = pixels[i + 1] = pixels[i + 2] = value;
                pixels[i + 3] = (byte)Math.Round(255 * (a + f));
            }
        Marshal.Copy(pixels, 0, bits, pixels.Length);

        var description = new PICTDESC { cbSizeofstruct = Marshal.SizeOf<PICTDESC>(), picType = PictureTypeBitmap, handle = bitmap };
        var iid = IPictureDispId;
        var hr = OleCreatePictureIndirect(ref description, ref iid, true, out var picture);
        if (hr < 0)
        {
            DeleteObject(bitmap);
            Marshal.ThrowExceptionForHR(hr);
        }
        return picture;
    }

    /// <summary>
    /// The screen's scale (1.25 at 125 %). Office can call add-ins in a DPI-unaware context,
    /// where Windows answers 96 dpi and then stretches the pictures, blurring them; so the
    /// question is asked as a DPI-aware thread.
    /// </summary>
    private static double DpiScale()
    {
        if (_scale is { } known)
            return known;

        uint asCalled = 0, dpi = 0;
        try
        {
            asCalled = GetDpiForSystem();
            var previous = SetThreadDpiAwarenessContext(DpiAwarenessContextSystemAware);
            try { dpi = GetDpiForSystem(); }
            finally
            {
                if (previous != IntPtr.Zero)
                    SetThreadDpiAwarenessContext(previous);
            }
        }
        catch (EntryPointNotFoundException) { } // before Windows 10 1607
        if (dpi == 0)
            dpi = asCalled > 0 ? asCalled : 96;
        _scale = dpi / 96.0;
        Log.Info($"Ribbon symbols drawn for {dpi} dpi (Word's thread reports {asCalled})");
        return _scale.Value;
    }

    /// <summary>
    /// Office's Dark Gray and Black themes (or "Use system setting" with Windows in dark mode)
    /// have dark menus, so the symbols are drawn light there.
    /// </summary>
    private static bool IsDarkTheme()
    {
        try
        {
            return Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common", "UI Theme", null) switch
            {
                3 or 4 => true,
                6 => Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1) is 0,
                _ => false,
            };
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static IntPtr CreateFont(int em, string face) =>
        CreateFontW(-em, 0, 0, 0, 400, 0, 0, 0, DefaultCharset, OutTrueTypePrecision, 0, AntialiasedQuality, 0, face);

    private static bool HasGlyph(IntPtr dc, char c)
    {
        var index = new ushort[1];
        return GetGlyphIndicesW(dc, c.ToString(), 1, index, MarkNonexistingGlyphs) == 1 && index[0] != 0xFFFF;
    }

    // ---- Win32 ----------------------------------------------------------------------------

    private const int Transparent = 1;
    private const uint TextAlignBaseline = 24;
    private const uint DefaultCharset = 1;
    private const uint OutTrueTypePrecision = 4;
    private const uint AntialiasedQuality = 4;
    private const uint MarkNonexistingGlyphs = 1;
    private const int PictureTypeBitmap = 1;
    private static readonly IntPtr DpiAwarenessContextSystemAware = new(-2);
    private static readonly Guid IPictureDispId = new("7BF80981-BF32-101A-8BBB-00AA00300CAB");

    [StructLayout(LayoutKind.Sequential)]
    private struct BITMAPINFOHEADER
    {
        public uint biSize;
        public int biWidth;
        public int biHeight;
        public ushort biPlanes;
        public ushort biBitCount;
        public uint biCompression;
        public uint biSizeImage;
        public int biXPelsPerMeter;
        public int biYPelsPerMeter;
        public uint biClrUsed;
        public uint biClrImportant;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PICTDESC
    {
        public int cbSizeofstruct;
        public int picType;
        public IntPtr handle;   // HBITMAP
        public IntPtr palette;  // HPALETTE, unused
    }

    [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr dc);
    [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
    [DllImport("gdi32.dll")] private static extern IntPtr CreateDIBSection(IntPtr dc, ref BITMAPINFOHEADER header, uint usage, out IntPtr bits, IntPtr section, uint offset);
    [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr gdiObject);
    [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr gdiObject);
    [DllImport("gdi32.dll")] private static extern uint SetTextColor(IntPtr dc, uint color);
    [DllImport("gdi32.dll")] private static extern int SetBkMode(IntPtr dc, int mode);
    [DllImport("gdi32.dll")] private static extern uint SetTextAlign(IntPtr dc, uint align);
    [DllImport("gdi32.dll")] private static extern bool GdiFlush();

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern bool TextOutW(IntPtr dc, int x, int y, string text, int length);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateFontW(int height, int width, int escapement, int orientation, int weight,
        uint italic, uint underline, uint strikeOut, uint charSet, uint outPrecision, uint clipPrecision,
        uint quality, uint pitchAndFamily, string face);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern uint GetGlyphIndicesW(IntPtr dc, string text, int length, [Out] ushort[] indices, uint flags);

    [DllImport("user32.dll")] private static extern uint GetDpiForSystem();
    [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    [DllImport("oleaut32.dll")]
    private static extern int OleCreatePictureIndirect(ref PICTDESC description, ref Guid iid, bool ownsHandle,
        [MarshalAs(UnmanagedType.IDispatch)] out object picture);
}
