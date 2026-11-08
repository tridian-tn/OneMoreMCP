using System.Xml.Linq;
using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>Covers the local title search used when OneMore's SearchTitles can't run.</summary>
public class TitleSearchTests
{
    private const string Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    [Fact]
    public void ParseTerms_keeps_quoted_phrases_and_drops_AND_and_the_sort_marker()
    {
        Assert.Equal(new[] { "road map", "Q3", "plan" }, TitleSearch.ParseTerms("> \"road map\" AND Q3 plan"));
    }

    [Fact]
    public void ParseTerms_keeps_a_quoted_AND()
    {
        Assert.Equal(new[] { "AND" }, TitleSearch.ParseTerms("\"AND\""));
    }

    [Fact]
    public void Matches_needs_every_term_in_any_order_ignoring_case()
    {
        var terms = TitleSearch.ParseTerms("plan ROAD");
        Assert.True(TitleSearch.Matches("Road trip plan", terms));
        Assert.False(TitleSearch.Matches("Road trip", terms));
    }

    [Fact]
    public void Matches_is_false_with_no_terms()
    {
        Assert.False(TitleSearch.Matches("Anything", Array.Empty<string>()));
    }

    [Fact]
    public void ListSections_walks_section_groups_across_notebooks_and_skips_recycle_bins()
    {
        var xml =
            $"<one:Notebooks xmlns:one=\"{Ns}\">" +
            "<one:Notebook name=\"Work\">" +
            "<one:Section name=\"Ideas\" />" +
            "<one:SectionGroup name=\"Archive\"><one:SectionGroup name=\"Old\"><one:Section name=\"2024\" /></one:SectionGroup></one:SectionGroup>" +
            "<one:SectionGroup name=\"OneNote_RecycleBin\" isRecycleBin=\"true\"><one:Section name=\"Deleted Pages\" isInRecycleBin=\"true\" /></one:SectionGroup>" +
            "</one:Notebook>" +
            "<one:Notebook name=\"Home\"><one:Section name=\"Lists\" /></one:Notebook>" +
            "</one:Notebooks>";

        Assert.Equal(
            new[]
            {
                new TitleSearch.SectionRef("Work", "Ideas"),
                new TitleSearch.SectionRef("Work", "Archive/Old/2024"),
                new TitleSearch.SectionRef("Home", "Lists"),
            },
            TitleSearch.ListSections(xml));
    }

    private const string TwoNotebooks =
        $"<one:Notebooks xmlns:one=\"{Ns}\">" +
        "<one:Notebook name=\"Work\"><one:Section name=\"Ideas\" /></one:Notebook>" +
        "<one:Notebook name=\"WorkArchive\"><one:Section name=\"2024\" /></one:Notebook>" +
        "<one:Notebook name=\"Home\"><one:Section name=\"Lists\" /></one:Notebook>" +
        "</one:Notebooks>";

    [Fact]
    public void ListSections_filter_matches_notebook_names_partially_ignoring_case_as_SearchTitles_does()
    {
        Assert.Equal(new[] { "Work", "WorkArchive" },
            TitleSearch.ListSections(TwoNotebooks, "work")!.Select(s => s.Notebook).ToArray());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("*")]
    public void ListSections_without_a_filter_keeps_every_notebook(string? filter)
    {
        Assert.Equal(3, TitleSearch.ListSections(TwoNotebooks, filter)!.Count);
    }

    [Fact]
    public void ListSections_is_null_when_no_notebook_matches_the_filter()
    {
        Assert.Null(TitleSearch.ListSections(TwoNotebooks, "Nope"));
        Assert.Null(TitleSearch.ListSections("", "Nope"));
    }

    [Fact]
    public void ListSections_is_empty_for_blank_output_and_throws_for_a_cli_message()
    {
        Assert.Empty(TitleSearch.ListSections("")!);
        var ex = Assert.Throws<ArgumentException>(() => TitleSearch.ListSections("Cannot connect to OneNote"));
        Assert.Contains("Cannot connect to OneNote", ex.Message);
    }

    [Fact]
    public void MatchPages_builds_paths_and_skips_recycled_pages()
    {
        var xml =
            $"<one:Section xmlns:one=\"{Ns}\" name=\"2024\">" +
            "<one:Page ID=\"{1}\" name=\"Roadmap\" lastModifiedTime=\"2026-01-02T03:04:05.000Z\" />" +
            "<one:Page ID=\"{2}\" name=\"Old roadmap\" isInRecycleBin=\"true\" />" +
            "<one:Page ID=\"{3}\" name=\"Budget\" />" +
            "</one:Section>";

        var pages = TitleSearch.MatchPages(xml, new TitleSearch.SectionRef("Work", "Archive/2024"), TitleSearch.ParseTerms("roadmap"));

        var page = Assert.Single(pages!);
        Assert.Equal("{1}", (string?)page.Attribute("id"));
        Assert.Equal("Work/Archive/2024/Roadmap", (string?)page.Attribute("path"));
        Assert.Equal("2026-01-02T03:04:05.000Z", (string?)page.Attribute("modified"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("Section not found: Work/Nope")]
    [InlineData("<one:Notebook xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" />")]
    public void MatchPages_is_null_when_the_output_isnt_a_section(string output)
    {
        Assert.Null(TitleSearch.MatchPages(output, new TitleSearch.SectionRef("Work", "Nope"), new[] { "x" }));
    }

    [Fact]
    public void BuildResults_renders_through_the_existing_markdown_projection()
    {
        var page = new XElement("Page", new XAttribute("id", "{1}"), new XAttribute("path", "Work/Ideas/Roadmap"));

        var xml = TitleSearch.BuildResults("roadmap", new[] { page }, skippedSections: 2);

        var root = XElement.Parse(xml);
        Assert.Equal("1", (string?)root.Attribute("count"));
        Assert.Equal("2", (string?)root.Attribute("skippedSections"));
        Assert.Equal("- Work/Ideas/Roadmap", OneNoteContent.SearchResultsToMarkdown(xml));
    }
}
