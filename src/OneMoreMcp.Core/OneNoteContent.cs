using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// Best-effort, read-only projections of OneNote XML into compact Markdown. OneNote's page and
/// hierarchy XML is verbose and namespaced; sending it verbatim to an LLM is token-heavy, so the
/// read tools summarise by default (raw XML remains available on request). The transforms are
/// lossy on formatting by design — they preserve structure and text, not styling.
/// </summary>
public static partial class OneNoteContent
{
    /// <summary>
    /// Renders a page's outline text as Markdown: the title as an H1, then each outline's content
    /// as (optionally nested) lines. Text runs are stripped of their inline HTML. Falls back to the
    /// raw input if it can't be parsed as a page.
    /// </summary>
    public static string PageToMarkdown(string pageXml)
    {
        XDocument doc;
        try { doc = XDocument.Parse(pageXml); }
        catch (System.Xml.XmlException) { return pageXml; }

        var page = doc.Root;
        if (page is null || page.Name != OneNoteSchema.Page) return pageXml;

        var sb = new StringBuilder();
        var title = page.Element(OneNoteSchema.Title);
        var titleText = title is null ? null : PlainText(AllRuns(title));
        if (!string.IsNullOrWhiteSpace(titleText))
            sb.Append("# ").AppendLine(titleText.Trim());

        foreach (var outline in page.Elements(OneNoteSchema.Outline))
        {
            foreach (var children in outline.Elements(OneNoteSchema.OEChildren))
                WriteOEChildren(children, sb, depth: 0);
        }

        var text = sb.ToString().TrimEnd();
        return text.Length == 0 ? "*(page has no text content)*" : text;
    }

    /// <summary>
    /// Renders GetHierarchy output as an indented tree of notebook / section / page names. Falls
    /// back to the raw input if it can't be parsed.
    /// </summary>
    public static string HierarchyToMarkdown(string hierarchyXml)
    {
        XDocument doc;
        try { doc = XDocument.Parse(hierarchyXml); }
        catch (System.Xml.XmlException) { return hierarchyXml; }
        if (doc.Root is null) return hierarchyXml;

        var sb = new StringBuilder();
        WriteHierarchy(doc.Root, sb, depth: -1); // root container itself isn't printed
        var text = sb.ToString().TrimEnd();
        return text.Length == 0 ? "*(no notebooks/sections/pages found)*" : text;
    }

    private static void WriteHierarchy(XElement element, StringBuilder sb, int depth)
    {
        var local = element.Name.LocalName;
        var isNode = local is "Notebook" or "SectionGroup" or "Section" or "Page";
        if (isNode)
        {
            var name = (string?)element.Attribute("name") ?? (string?)element.Attribute("ID") ?? local;
            sb.Append(new string(' ', Math.Max(0, depth) * 2)).Append("- ").AppendLine(name.Trim());
        }

        foreach (var child in element.Elements())
            WriteHierarchy(child, sb, depth + 1);
    }

    private static void WriteOEChildren(XElement children, StringBuilder sb, int depth)
    {
        foreach (var oe in children.Elements(OneNoteSchema.OE))
        {
            var line = PlainText(oe.Elements(OneNoteSchema.T));
            if (!string.IsNullOrWhiteSpace(line))
                sb.Append(new string(' ', depth * 2)).Append("- ").AppendLine(line.Trim());

            foreach (var nested in oe.Elements(OneNoteSchema.OEChildren))
                WriteOEChildren(nested, sb, depth + 1);
        }
    }

    private static IEnumerable<XElement> AllRuns(XElement container) =>
        container.Descendants(OneNoteSchema.T);

    /// <summary>Concatenates the text of <c>one:T</c> runs and strips their inline HTML to plain text.</summary>
    private static string PlainText(IEnumerable<XElement> runs)
    {
        var sb = new StringBuilder();
        foreach (var run in runs) sb.Append(run.Value);
        return StripHtml(sb.ToString());
    }

    /// <summary>Turns a OneNote text run's inner HTML into plain text (breaks → spaces, tags removed, entities decoded).</summary>
    public static string StripHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return string.Empty;
        var withBreaks = BreakTag().Replace(html, " ");
        var noTags = TagRun().Replace(withBreaks, string.Empty);
        return WebUtility.HtmlDecode(noTags).Trim();
    }

    [GeneratedRegex(@"<br\s*/?>", RegexOptions.IgnoreCase)]
    private static partial Regex BreakTag();

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex TagRun();
}
