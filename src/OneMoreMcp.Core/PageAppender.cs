using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>The authoring format of text handed to <see cref="PageAppender"/>.</summary>
public enum AppendFormat
{
    /// <summary>Literal text; special characters are HTML-encoded, one paragraph per line.</summary>
    Plain,

    /// <summary>A small Markdown subset (headings, bullets, <c>**bold**</c>, <c>*italic*</c>, <c>`code`</c>).</summary>
    Markdown,

    /// <summary>Raw HTML placed into the OneNote text run verbatim (the caller owns validity).</summary>
    Html,
}

/// <summary>
/// Appends text to a OneNote page by editing the page's XML: it inserts new <c>one:OE</c> paragraphs
/// into the last outline's <c>one:OEChildren</c> (creating an outline if the page has none) and leaves
/// everything else untouched. This is the engine behind the ungated <c>append_to_page</c> tool — the
/// page's existing content is fetched, mutated locally, and written back, so it never becomes LLM
/// tokens and can never be overwritten, only added to.
/// </summary>
public static partial class PageAppender
{
    /// <summary>
    /// Returns <paramref name="pageXml"/> with <paramref name="text"/> appended as one or more
    /// paragraphs. The input must be a <c>one:Page</c> document (as GetPage emits). Existing outlines
    /// and their content are preserved; only new <c>one:OE</c> nodes are added.
    /// </summary>
    public static string Append(string pageXml, string text, AppendFormat format)
    {
        if (string.IsNullOrEmpty(pageXml))
            throw new ArgumentException("The page XML is empty.", nameof(pageXml));
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("There is no text to append.", nameof(text));

        XDocument doc;
        try { doc = XDocument.Parse(pageXml, LoadOptions.PreserveWhitespace); }
        catch (XmlException ex) { throw new ArgumentException("The page XML could not be parsed.", nameof(pageXml), ex); }

        var page = doc.Root;
        if (page is null || page.Name != OneNoteSchema.Page)
            throw new ArgumentException("The provided XML is not a OneNote page (expected a one:Page root).", nameof(pageXml));

        var paragraphs = BuildParagraphs(text ?? string.Empty, format).ToList();
        if (paragraphs.Count == 0)
            throw new ArgumentException("There is no text to append.", nameof(text));

        var target = AppendTarget(page);
        foreach (var oe in paragraphs) target.Add(oe);

        return Serialize(doc);
    }

    /// <summary>
    /// Finds (or creates) the <c>one:OEChildren</c> to append into: the last outline's last child
    /// container. A page with no outline gets a fresh outline + container added at the end.
    /// </summary>
    private static XElement AppendTarget(XElement page)
    {
        var outline = page.Elements(OneNoteSchema.Outline).LastOrDefault();
        if (outline is null)
        {
            outline = new XElement(OneNoteSchema.Outline, new XElement(OneNoteSchema.OEChildren));
            page.Add(outline);
        }

        var children = outline.Elements(OneNoteSchema.OEChildren).LastOrDefault();
        if (children is null)
        {
            children = new XElement(OneNoteSchema.OEChildren);
            outline.Add(children);
        }
        return children;
    }

    private static IEnumerable<XElement> BuildParagraphs(string text, AppendFormat format)
    {
        // Normalise line endings, then one paragraph (one:OE) per line — mirroring how OneNote
        // models separate paragraphs. Empty lines become empty paragraphs (visual spacing).
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        foreach (var line in lines)
            yield return Paragraph(RunHtml(line, format));
    }

    private static XElement Paragraph(string innerHtml) =>
        new(OneNoteSchema.OE, new XElement(OneNoteSchema.T, new XCData(innerHtml)));

    /// <summary>Converts a single source line into the inner HTML of a <c>one:T</c> run.</summary>
    private static string RunHtml(string line, AppendFormat format) => format switch
    {
        AppendFormat.Html => line,                       // caller-supplied HTML, verbatim
        AppendFormat.Markdown => MarkdownLine(line),
        _ => WebUtility.HtmlEncode(line),                // Plain
    };

    /// <summary>
    /// A deliberately small Markdown-to-HTML line converter: <c>#…######</c> headings (bold), <c>-</c>/<c>*</c>
    /// bullets (bullet glyph + text), and inline <c>**bold**</c>, <c>*italic*</c>/<c>_italic_</c>, <c>`code`</c>.
    /// Anything else is treated as literal text. Encoding happens first so user text can't inject markup.
    /// </summary>
    private static string MarkdownLine(string line)
    {
        var trimmed = line.TrimStart();
        var indent = line.Length - trimmed.Length;

        string? prefix = null;
        var heading = HeadingMarker().Match(trimmed);
        if (heading.Success)
        {
            trimmed = trimmed[heading.Length..];
            prefix = "heading";
        }
        else if (BulletMarker().IsMatch(trimmed))
        {
            trimmed = BulletMarker().Replace(trimmed, string.Empty);
            prefix = "bullet";
        }

        var html = Inline(WebUtility.HtmlEncode(trimmed));
        html = prefix switch
        {
            "heading" => $"<span style='font-weight:bold'>{html}</span>",
            "bullet" => "• " + html,
            _ => html,
        };
        return new string(' ', indent) + html;
    }

    private static string Inline(string encoded)
    {
        encoded = BoldToken().Replace(encoded, "<span style='font-weight:bold'>$1</span>");
        encoded = ItalicStar().Replace(encoded, "<span style='font-style:italic'>$1</span>");
        encoded = ItalicUnderscore().Replace(encoded, "<span style='font-style:italic'>$1</span>");
        encoded = CodeToken().Replace(encoded, "<span style='font-family:Consolas'>$1</span>");
        return encoded;
    }

    private static string Serialize(XDocument doc)
    {
        var settings = new XmlWriterSettings { OmitXmlDeclaration = false, Indent = false };
        var sb = new StringBuilder();
        using var writer = XmlWriter.Create(sb, settings);
        doc.Save(writer);
        writer.Flush();
        return sb.ToString();
    }

    [GeneratedRegex(@"^#{1,6}\s+")]
    private static partial Regex HeadingMarker();

    [GeneratedRegex(@"^[-*]\s+")]
    private static partial Regex BulletMarker();

    [GeneratedRegex(@"\*\*(.+?)\*\*")]
    private static partial Regex BoldToken();

    [GeneratedRegex(@"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])")]
    private static partial Regex ItalicStar();

    [GeneratedRegex(@"(?<!\w)_(?!\s)(.+?)(?<!\s)_(?!\w)")]
    private static partial Regex ItalicUnderscore();

    [GeneratedRegex("`(.+?)`")]
    private static partial Regex CodeToken();
}
