using System.Xml;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// Compares two copies of a OneNote page to tell whether its content really differs. Comparing the raw
/// XML can't answer that: OneNote moves <c>lastModifiedTime</c> even for a write it discards, and the
/// selection and view state shift as someone uses the page. This only reads its inputs, so the XML that's
/// written back to a page keeps its <c>omHash</c>.
/// </summary>
public static class PageXmlComparer
{
    // omHash belongs here because OneMore computes it over each top-level element's serialized XML, these
    // other attributes included, so it moves whenever they do. It adds nothing to the comparison either,
    // since a real change shows up in the content it's derived from.
    private static readonly HashSet<string> VolatileAttributes = new()
    {
        "lastModifiedTime",
        "selected",
        "isCurrentlyViewed",
        "omHash",
    };

    /// <summary>
    /// Checks whether two page documents hold the same content, ignoring volatile attributes and how the
    /// XML happens to be written: formatting whitespace, CDATA versus plain text, and namespace prefixes.
    /// Text that can't be parsed only matches an identical string.
    /// </summary>
    /// <param name="first">One copy of the page XML</param>
    /// <param name="second">The copy to compare it with</param>
    /// <returns>True if nothing but volatile attributes or the way the XML's written differs</returns>
    public static bool HasSameContent(string first, string second)
    {
        if (string.Equals(first, second, StringComparison.Ordinal)) return true;

        var firstContent = ParseContent(first);
        var secondContent = ParseContent(second);
        return firstContent is not null && secondContent is not null && XNode.DeepEquals(firstContent, secondContent);
    }

    /// <summary>Parses page XML and reduces it to what can differ between copies of different content.</summary>
    /// <param name="xml">The page XML</param>
    /// <returns>The reduced document, or null if the text isn't XML</returns>
    private static XDocument? ParseContent(string xml)
    {
        // Default parsing is deliberate: OneMore's PutPage reads the XML the same way, so whitespace-only
        // plain text counts as empty here because empty is what gets written. Keeping it would make a
        // space sent into an empty run look like a requested change that never landed.
        XDocument doc;
        try { doc = XDocument.Parse(xml); }
        catch (XmlException) { return null; }

        // Dropping namespace declarations only loses the prefixes, since every name keeps its namespace.
        doc.Descendants().Attributes()
            .Where(a => a.IsNamespaceDeclaration || VolatileAttributes.Contains(a.Name.LocalName))
            .Remove();

        // XNode.DeepEquals treats CDATA and plain text holding the same characters as different.
        foreach (var cdata in doc.DescendantNodes().OfType<XCData>().ToList()) cdata.ReplaceWith(new XText(cdata.Value));

        return doc;
    }
}
