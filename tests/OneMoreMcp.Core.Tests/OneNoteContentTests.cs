using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>Covers the read-only XML → Markdown projections.</summary>
public class OneNoteContentTests
{
    private const string Page =
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">" +
        "<one:Title><one:OE><one:T><![CDATA[My Page]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[<span style='font-weight:bold'>First</span>]]></one:T>" +
        "<one:OEChildren><one:OE><one:T><![CDATA[Nested]]></one:T></one:OE></one:OEChildren>" +
        "</one:OE>" +
        "<one:OE><one:T><![CDATA[Second &amp; last]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline>" +
        "</one:Page>";

    [Fact]
    public void PageToMarkdown_renders_title_as_h1()
    {
        var md = OneNoteContent.PageToMarkdown(Page);
        Assert.StartsWith("# My Page", md);
    }

    [Fact]
    public void PageToMarkdown_strips_inline_html_and_indents_nested_content()
    {
        var md = OneNoteContent.PageToMarkdown(Page);
        Assert.Contains("- First", md);
        Assert.Contains("  - Nested", md);        // nested one level deeper
        Assert.Contains("- Second & last", md);   // entity decoded
        Assert.DoesNotContain("font-weight", md); // inline HTML removed
    }

    [Fact]
    public void PageToMarkdown_returns_raw_input_when_not_a_page()
    {
        Assert.Equal("<other/>", OneNoteContent.PageToMarkdown("<other/>"));
    }

    [Fact]
    public void HierarchyToMarkdown_renders_an_indented_tree_of_names()
    {
        const string hierarchy =
            "<one:Notebooks xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">" +
            "<one:Notebook name=\"Work\"><one:Section name=\"Ideas\">" +
            "<one:Page name=\"Alpha\"/></one:Section></one:Notebook>" +
            "</one:Notebooks>";
        var md = OneNoteContent.HierarchyToMarkdown(hierarchy);
        Assert.Contains("- Work", md);
        Assert.Contains("  - Ideas", md);
        Assert.Contains("    - Alpha", md);
    }

    [Fact]
    public void SearchResultsToMarkdown_lists_page_paths()
    {
        const string results =
            "<Results query=\"Lorem\" count=\"2\">" +
            "<Page id=\"{1}\" path=\"testnotebook/New Section 1/Lorem ipsum\" />" +
            "<Page id=\"{2}\" path=\"Work/Ideas/Roadmap\" /></Results>";
        var md = OneNoteContent.SearchResultsToMarkdown(results);
        Assert.Equal("- testnotebook/New Section 1/Lorem ipsum\n- Work/Ideas/Roadmap", md.Replace("\r\n", "\n"));
    }

    [Fact]
    public void SearchResultsToMarkdown_reports_no_matches_for_empty_results()
    {
        Assert.Equal("*(no matching pages)*",
            OneNoteContent.SearchResultsToMarkdown("<Results query=\"x\" count=\"0\" />"));
    }

    [Fact]
    public void StripHtml_turns_breaks_into_spaces_and_decodes_entities()
    {
        Assert.Equal("a b", OneNoteContent.StripHtml("a<br/>b"));
        Assert.Equal("<tag>", OneNoteContent.StripHtml("&lt;tag&gt;"));
    }
}
