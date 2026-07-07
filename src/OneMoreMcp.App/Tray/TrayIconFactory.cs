using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace OneMoreMcp.App.Tray;

/// <summary>
/// Single source of truth for the app artwork: white note lines on a purple rounded square (evoking
/// OneNote). <see cref="Create"/> produces the tray <see cref="Icon"/>; <see cref="WriteIco"/> emits
/// the multi-resolution <c>App.ico</c> used as the executable icon (run with <c>--write-icon</c>).
/// </summary>
internal static class TrayIconFactory
{
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr handle);

    /// <summary>Sizes baked into the generated .ico (Explorer / taskbar / Alt-Tab at various DPI).</summary>
    private static readonly int[] IconSizes = { 16, 20, 24, 32, 48, 64, 128, 256 };

    /// <summary>Builds the tray icon at the current small-icon size (DPI-aware, crisp).</summary>
    public static Icon Create()
    {
        // Small-icon size is normally square; take the larger dimension defensively so the
        // rendered (square) bitmap is never smaller than what the shell expects.
        var small = SystemInformation.SmallIconSize;
        var size = Math.Max(16, Math.Max(small.Width, small.Height));
        using var bmp = RenderBitmap(size);
        var hicon = bmp.GetHicon();
        try
        {
            using var temp = Icon.FromHandle(hicon);
            return (Icon)temp.Clone();
        }
        finally
        {
            DestroyIcon(hicon);
        }
    }

    /// <summary>Writes a multi-resolution Windows .ico to <paramref name="path"/>.</summary>
    public static void WriteIco(string path) => File.WriteAllBytes(path, BuildIco(IconSizes));

    /// <summary>Renders the artwork to a fresh <see cref="Bitmap"/> of the given square size (the caller owns it).</summary>
    public static Bitmap CreateBitmap(int size)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(size);
        return RenderBitmap(size);
    }

    private static byte[] BuildIco(int[] sizes)
    {
        var ordered = sizes.Distinct().OrderBy(s => s).ToArray();
        var frames = ordered.Select(PngFrame).ToArray();

        using var ms = new MemoryStream();
        using var w = new BinaryWriter(ms);

        // ICONDIR
        w.Write((ushort)0);                 // reserved
        w.Write((ushort)1);                 // type: icon
        w.Write((ushort)frames.Length);     // image count

        var offset = 6 + 16 * frames.Length;
        for (var i = 0; i < frames.Length; i++)
        {
            var dim = ordered[i] >= 256 ? 0 : ordered[i]; // 0 means 256 in the .ico format
            w.Write((byte)dim);             // width
            w.Write((byte)dim);             // height
            w.Write((byte)0);               // palette count
            w.Write((byte)0);               // reserved
            w.Write((ushort)1);             // colour planes
            w.Write((ushort)32);            // bits per pixel
            w.Write(frames[i].Length);      // bytes of image data
            w.Write(offset);                // offset to image data
            offset += frames[i].Length;
        }

        foreach (var frame in frames) w.Write(frame);
        w.Flush();
        return ms.ToArray();
    }

    private static byte[] PngFrame(int size)
    {
        using var bmp = RenderBitmap(size);
        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png); // PNG-compressed frames (Vista+; required for 256px)
        return ms.ToArray();
    }

    private static Bitmap RenderBitmap(int size)
    {
        var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.Clear(Color.Transparent);

        var f = size / 16f; // design is authored at 16px and scaled up

        // OneNote-ish purple rounded square.
        using var fill = new SolidBrush(Color.FromArgb(0x7B, 0x30, 0x89));
        g.FillRoundedRectangle(fill, new RectangleF(0, 0, size - 1, size - 1), 3f * f);

        // Three white "note" lines of decreasing length.
        using var pen = new Pen(Color.White, Math.Max(1.4f, 1.7f * f))
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
        };
        g.DrawLine(pen, 4.0f * f, 5.0f * f, 12.0f * f, 5.0f * f);
        g.DrawLine(pen, 4.0f * f, 8.0f * f, 12.0f * f, 8.0f * f);
        g.DrawLine(pen, 4.0f * f, 11.0f * f, 9.0f * f, 11.0f * f);

        return bmp;
    }
}

internal static class GraphicsExtensions
{
    public static void FillRoundedRectangle(this Graphics g, Brush brush, RectangleF bounds, float radius)
    {
        using var path = new GraphicsPath();
        var d = radius * 2;
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        g.FillPath(brush, path);
    }
}
