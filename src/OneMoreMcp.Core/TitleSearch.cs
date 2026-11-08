using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// A local stand-in for OneMore's <c>SearchTitles</c>, for when the CLI can't run it. It matches page
/// titles from <c>GetHierarchy</c> output itself: every word or quoted phrase in the query has to appear
/// in the title, ignoring case and order. OneMore's richer syntax (OR, NOT, parentheses, wildcards,
/// <c>&gt;</c> sorting, <c>#hashtag</c> filters) isn't supported here.
/// </summary>
public static partial class TitleSearch
{
    /// <summary>A section to read pages from: its notebook, and its '/'-delimited path within that notebook.</summary>
    public sealed record SectionRef(string Notebook, string Path);

    /// <summary>
    /// The sections to search, and the notebooks left out because an earlier open notebook has the same
    /// name. The CLI addresses a section by notebook name and takes the first notebook of that name, so a
    /// later one can't be reached; reading it would just read the first one again.
    /// </summary>
    public sealed record SectionList(IReadOnlyList<SectionRef> Sections, IReadOnlyList<string> UnreachableNotebooks);

    /// <summary>
    /// Lists every section in <c>GetHierarchy</c> output, including sections nested in section groups,
    /// for the notebooks a filter picks. Recycle bins are skipped, and so is any notebook that shares its
    /// name with an earlier one, which is reported in <see cref="SectionList.UnreachableNotebooks"/>.
    /// </summary>
    /// <param name="hierarchyXml">GetHierarchy output without <c>--section</c></param>
    /// <param name="notebookFilter">
    /// Picks notebooks the way OneMore's SearchTitles does: <c>\</c> for the one open in OneNote, otherwise
    /// those whose name contains this, ignoring case (so "Work" also picks "WorkArchive"). Null, blank or
    /// <c>*</c> keeps them all
    /// </param>
    /// <returns>The sections in hierarchy order, or null when a filter is given and no notebook matches it</returns>
    /// <exception cref="ArgumentException">The output isn't XML, which means the CLI printed an error</exception>
    public static SectionList? ListSections(string hierarchyXml, string? notebookFilter = null)
    {
        var filtered = !string.IsNullOrWhiteSpace(notebookFilter) && notebookFilter != "*";
        if (string.IsNullOrWhiteSpace(hierarchyXml))
        {
            return filtered ? null : new SectionList(Array.Empty<SectionRef>(), Array.Empty<string>());
        }

        XDocument doc;
        try { doc = XDocument.Parse(hierarchyXml); }
        catch (System.Xml.XmlException) { throw new ArgumentException(hierarchyXml.Trim()); }

        var sections = new List<SectionRef>();
        var unreachable = new List<string>();
        // The CLI matches notebook names this way when it resolves --notebook.
        var seen = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        var matched = 0;
        foreach (var notebook in doc.Root!.DescendantsAndSelf(OneNoteSchema.Ns + "Notebook"))
        {
            var name = (string?)notebook.Attribute("name") ?? "";
            // Judged across every notebook, not just the picked ones: --notebook goes to the first of a
            // name even when that one isn't picked.
            var reachable = seen.Add(name);
            if (filtered && !Picks(notebookFilter!, notebook, name)) continue;

            matched++;
            if (!reachable)
            {
                unreachable.Add(name);
                continue;
            }
            CollectSections(notebook, name, prefix: "", sections);
        }
        return filtered && matched == 0 ? null : new SectionList(sections, unreachable);
    }

    private static bool Picks(string filter, XElement notebook, string name) =>
        filter == @"\"
            ? (string?)notebook.Attribute("isCurrentlyViewed") == "true"
            : name.Contains(filter, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Splits a query into the words and quoted phrases a title has to contain. A bare <c>AND</c> is
    /// dropped, since every term has to match anyway, and so is a bare <c>&gt;</c> (OneMore's sort marker).
    /// </summary>
    public static IReadOnlyList<string> ParseTerms(string query)
    {
        var terms = new List<string>();
        foreach (Match m in TermPattern().Matches(query ?? ""))
        {
            var quoted = m.Groups[1].Success;
            var term = (quoted ? m.Groups[1].Value : m.Groups[2].Value).Trim();
            if (term.Length == 0) continue;
            if (!quoted && (term == "AND" || term == ">")) continue;
            terms.Add(term);
        }
        return terms;
    }

    /// <summary>Whether a title contains every term, ignoring case and order.</summary>
    public static bool Matches(string title, IReadOnlyList<string> terms) =>
        terms.Count > 0 && terms.All(t => title.Contains(t, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Finds the pages in one section's <c>GetHierarchy</c> output whose titles match, as
    /// <c>SearchTitles</c>-style <c>&lt;Page&gt;</c> elements carrying <c>id</c>, <c>path</c> and
    /// <c>modified</c>. Pages in the recycle bin are skipped.
    /// </summary>
    /// <param name="sectionXml">GetHierarchy output for <paramref name="section"/></param>
    /// <param name="section">The section the output is for, used to build each page's path</param>
    /// <param name="terms">The terms from <see cref="ParseTerms"/></param>
    /// <returns>The matching pages, or null when the output isn't a section (the CLI couldn't read it)</returns>
    public static IReadOnlyList<XElement>? MatchPages(string sectionXml, SectionRef section, IReadOnlyList<string> terms)
    {
        if (string.IsNullOrWhiteSpace(sectionXml)) return null;

        XDocument doc;
        try { doc = XDocument.Parse(sectionXml); }
        catch (System.Xml.XmlException) { return null; }
        if (doc.Root?.Name != OneNoteSchema.Ns + "Section") return null;

        return doc.Root.Elements(OneNoteSchema.Page)
            .Where(p => p.Attribute("isInRecycleBin") is null)
            .Select(p => (Page: p, Name: (string?)p.Attribute("name") ?? ""))
            .Where(p => Matches(p.Name, terms))
            .Select(p => new XElement("Page",
                new XAttribute("id", (string?)p.Page.Attribute("ID") ?? ""),
                new XAttribute("path", $"{section.Notebook}/{section.Path}/{p.Name}"),
                new XAttribute("modified", (string?)p.Page.Attribute("lastModifiedTime") ?? "")))
            .ToList();
    }

    /// <summary>
    /// Wraps matched pages in the <c>&lt;Results&gt;</c> document <c>SearchTitles</c> produces, marked
    /// <c>source="local"</c> so a caller reading the XML can tell the fallback ran.
    /// </summary>
    /// <param name="query">The original query, echoed as SearchTitles does</param>
    /// <param name="pages">The pages from <see cref="MatchPages"/></param>
    /// <param name="skippedSections">How many sections couldn't be read</param>
    /// <param name="unreachableNotebooks">Notebooks left out because they share a name with an earlier one</param>
    public static string BuildResults(
        string query, IReadOnlyCollection<XElement> pages, int skippedSections, IReadOnlyList<string>? unreachableNotebooks = null)
    {
        var root = new XElement("Results",
            new XAttribute("query", query),
            new XAttribute("count", pages.Count),
            new XAttribute("source", "local"));
        if (skippedSections > 0) root.Add(new XAttribute("skippedSections", skippedSections));
        if (unreachableNotebooks is { Count: > 0 })
        {
            root.Add(new XAttribute("unreachableNotebooks", string.Join("; ", unreachableNotebooks)));
        }
        root.Add(pages);
        return root.ToString();
    }

    private static void CollectSections(XElement parent, string notebook, string prefix, List<SectionRef> sections)
    {
        foreach (var child in parent.Elements())
        {
            if (IsRecycleBin(child)) continue;

            var name = (string?)child.Attribute("name") ?? "";
            if (child.Name == OneNoteSchema.Ns + "Section")
            {
                sections.Add(new SectionRef(notebook, prefix + name));
            }
            else if (child.Name == OneNoteSchema.Ns + "SectionGroup")
            {
                CollectSections(child, notebook, prefix + name + "/", sections);
            }
        }
    }

    private static bool IsRecycleBin(XElement e) =>
        e.Attribute("isRecycleBin") is not null || e.Attribute("isInRecycleBin") is not null;

    // A double-quoted phrase, or a run of non-space characters.
    [GeneratedRegex("\"([^\"]*)\"|(\\S+)")]
    private static partial Regex TermPattern();
}
