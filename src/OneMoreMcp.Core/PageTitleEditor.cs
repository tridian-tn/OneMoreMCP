using System.Xml;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// Sets the title text on a OneNote page document. Creating a page is a two-step operation — the CLI
/// makes an untitled shell, then the title is applied by writing the page back — so this edits the
/// <c>one:Title</c> that OneNote put on the new page rather than building one from scratch.
/// </summary>
public static class PageTitleEditor
{
    /// <summary>
    /// Returns <paramref name="pageXml"/> with its title text replaced by <paramref name="title"/>,
    /// leaving every other element (and the page's IDs) untouched.
    /// </summary>
    /// <param name="pageXml">A <c>one:Page</c> document, as GetPage emits</param>
    /// <param name="title">The title text to set</param>
    /// <returns>The page XML with the new title applied</returns>
    public static string SetTitle(string pageXml, string title)
    {
        if (string.IsNullOrWhiteSpace(pageXml))
            throw new ArgumentException("The page XML is empty.", nameof(pageXml));
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("The title is empty.", nameof(title));

        XDocument doc;
        try { doc = XDocument.Parse(pageXml); }
        catch (XmlException ex) { throw new ArgumentException("The page XML could not be parsed.", nameof(pageXml), ex); }

        var page = doc.Root;
        if (page is null || page.Name != OneNoteSchema.Page)
            throw new ArgumentException("The provided XML is not a OneNote page (expected a one:Page root).", nameof(pageXml));

        // A page created by the CLI already carries an empty one:Title; reuse its OE so the objectID
        // OneNote assigned is preserved, since an update only applies to objects OneNote can match.
        var titleElement = page.Element(OneNoteSchema.Title);
        if (titleElement is null)
        {
            titleElement = new XElement(OneNoteSchema.Title,
                new XElement(OneNoteSchema.OE, new XElement(OneNoteSchema.T, new XCData(title))));
            page.AddFirst(titleElement);
            return doc.ToString(SaveOptions.DisableFormatting);
        }

        var oe = titleElement.Element(OneNoteSchema.OE);
        if (oe is null)
        {
            titleElement.Add(new XElement(OneNoteSchema.OE, new XElement(OneNoteSchema.T, new XCData(title))));
            return doc.ToString(SaveOptions.DisableFormatting);
        }

        // OneNote splits a title across several runs; collapse them into one carrying the new text.
        var runs = oe.Elements(OneNoteSchema.T).ToList();
        if (runs.Count == 0)
        {
            oe.Add(new XElement(OneNoteSchema.T, new XCData(title)));
        }
        else
        {
            runs[0].ReplaceNodes(new XCData(title));
            foreach (var extra in runs.Skip(1)) extra.Remove();
        }

        return doc.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Reads a page's title text, or null when the document has no title runs.</summary>
    /// <param name="pageXml">A <c>one:Page</c> document</param>
    /// <returns>The title text, or null if absent</returns>
    public static string? GetTitle(string pageXml)
    {
        XDocument doc;
        try { doc = XDocument.Parse(pageXml); }
        catch (XmlException) { return null; }

        var title = doc.Root?.Element(OneNoteSchema.Title);
        if (title is null) return null;

        var text = string.Concat(title.Descendants(OneNoteSchema.T).Select(t => t.Value));
        return string.IsNullOrWhiteSpace(text) ? null : OneNoteContent.StripHtml(text);
    }

    /// <summary>
    /// Builds the minimal page document handed to the CLI to create a page. The content never lands
    /// (OneNote ignores outlines added to a page that has none), but the CLI requires a schema-valid
    /// <c>one:Page</c> before it will create the shell.
    /// </summary>
    /// <param name="title">The intended page title</param>
    /// <returns>A schema-valid one:Page document</returns>
    public static string NewPageXml(string title)
    {
        var page = new XElement(OneNoteSchema.Page,
            new XAttribute(XNamespace.Xmlns + "one", OneNoteSchema.Ns.NamespaceName),
            new XElement(OneNoteSchema.Title,
                new XElement(OneNoteSchema.OE, new XElement(OneNoteSchema.T, new XCData(title)))));

        return new XDocument(page).ToString(SaveOptions.DisableFormatting);
    }
}
