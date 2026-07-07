using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// Prepares OneNote page XML for writing back via <c>PutPage</c>. <c>GetPage</c> emits some attributes
/// that OneNote recomputes and that PutPage's schema validation rejects (notably <c>omHash</c>), so they
/// must be stripped before a round-trip — otherwise PutPage fails schema validation yet still exits 0,
/// silently writing nothing.
/// </summary>
public static class OneNotePage
{
    // Attributes GetPage emits that PutPage's schema won't accept. OneNote recomputes them on save.
    private static readonly HashSet<string> VolatileAttributes = new(StringComparer.Ordinal) { "omHash" };

    /// <summary>Returns the page XML with write-rejected volatile attributes removed, ready for PutPage --infile.</summary>
    public static string PrepareForPut(string pageXml)
    {
        if (string.IsNullOrEmpty(pageXml))
            throw new ArgumentException("The page XML is empty.", nameof(pageXml));

        XDocument doc;
        try { doc = XDocument.Parse(pageXml, LoadOptions.PreserveWhitespace); }
        catch (XmlException ex) { throw new ArgumentException("The page XML could not be parsed.", nameof(pageXml), ex); }

        foreach (var element in doc.Descendants())
        {
            var drop = element.Attributes().Where(a => VolatileAttributes.Contains(a.Name.LocalName)).ToList();
            foreach (var attribute in drop) attribute.Remove();
        }

        var settings = new XmlWriterSettings { OmitXmlDeclaration = false, Indent = false };
        var sb = new StringBuilder();
        using var writer = XmlWriter.Create(sb, settings);
        doc.Save(writer);
        writer.Flush();
        return sb.ToString();
    }
}
