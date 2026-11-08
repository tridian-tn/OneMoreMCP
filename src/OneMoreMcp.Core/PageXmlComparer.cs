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
    /// XML happens to be written: formatting whitespace, CDATA versus plain text, namespace prefixes,
    /// attribute order, and comments. Text that can't be parsed only matches an identical string.
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
    /// <returns>The reduced root element, or null if the text isn't XML</returns>
    private static XElement? ParseContent(string xml)
    {
        // This parses the way OneMore's PutPage does, deliberately, so it sees what would be written. Anything
        // outside the root is dropped, and whitespace-only plain text counts as empty. Keeping that whitespace
        // would make a space sent into an empty run look like a requested change that never landed.
        XElement root;
        try { root = XElement.Parse(xml); }
        catch (XmlException) { return null; }

        // OneNote ignores these inside the page too. A comment in a run's HTML is CDATA, so it isn't touched.
        root.DescendantNodes().Where(n => n is XComment or XProcessingInstruction).Remove();

        // Dropping namespace declarations only loses the prefixes, since every name keeps its namespace.
        root.DescendantsAndSelf().Attributes()
            .Where(a => a.IsNamespaceDeclaration || VolatileAttributes.Contains(a.Name.LocalName))
            .Remove();

        // XNode.DeepEquals is sensitive to attribute order, which means nothing in XML.
        foreach (var element in root.DescendantsAndSelf())
        {
            element.ReplaceAttributes(element.Attributes().OrderBy(a => a.Name.ToString(), StringComparer.Ordinal).ToList());
        }

        // It also treats CDATA and plain text holding the same characters as different.
        foreach (var cdata in root.DescendantNodes().OfType<XCData>().ToList()) cdata.ReplaceWith(new XText(cdata.Value));

        return root;
    }
}
