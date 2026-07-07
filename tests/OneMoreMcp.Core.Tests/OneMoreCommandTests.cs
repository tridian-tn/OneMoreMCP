using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>Asserts the exact argument vectors the command factories produce.</summary>
public class OneMoreCommandTests
{
    [Fact]
    public void GetHierarchy_includes_only_supplied_options()
    {
        var argv = OneMoreCommand.GetHierarchy(notebook: "Work", section: null, books: true).Build();
        Assert.Equal(new[] { "GetHierarchy", "--notebook", "Work", "--books" }, argv);
    }

    [Fact]
    public void GetHierarchy_with_no_options_is_just_the_verb()
    {
        Assert.Equal(new[] { "GetHierarchy" }, OneMoreCommand.GetHierarchy().Build());
    }

    [Fact]
    public void GetPage_current_omits_notebook_section_and_page()
    {
        var argv = OneMoreCommand.GetPage(notebook: null, section: null, page: null, current: true).Build();
        Assert.Equal(new[] { "GetPage", "--current" }, argv);
    }

    [Fact]
    public void GetPage_by_name_includes_notebook_and_keeps_spaces_unquoted()
    {
        var argv = OneMoreCommand.GetPage(notebook: "My Book", section: "My Section", page: "Meeting Notes", current: false).Build();
        Assert.Equal(new[] { "GetPage", "--notebook", "My Book", "--section", "My Section", "--page", "Meeting Notes" }, argv);
    }

    [Fact]
    public void PutPage_forces_by_default_and_carries_the_infile()
    {
        var argv = OneMoreCommand.PutPage(notebook: "B", section: "S", page: null, infile: @"C:\t\p.xml").Build();
        Assert.Equal(new[] { "PutPage", "--notebook", "B", "--section", "S", "--infile", @"C:\t\p.xml", "--force" }, argv);
    }

    [Fact]
    public void SearchHashtags_allTags_adds_the_switch()
    {
        Assert.Equal(
            new[] { "SearchHashtags", "--query", "#a #b", "--allTags" },
            OneMoreCommand.SearchHashtags("#a #b", allTags: true).Build());
    }

    [Fact]
    public void Export_includes_format_and_optional_pageId()
    {
        var argv = OneMoreCommand.Export(outpath: @"C:\out", format: "PDF", pageId: "{123}").Build();
        Assert.Equal(new[] { "Export", "--outpath", @"C:\out", "--format", "PDF", "--pageId", "{123}" }, argv);
    }

    [Fact]
    public void ToString_quotes_only_arguments_with_spaces()
    {
        var cmd = OneMoreCommand.GetPage(notebook: "Book", section: "My Section", page: "Notes", current: false);
        Assert.Equal("GetPage --notebook Book --section \"My Section\" --page Notes", cmd.ToString());
    }

    [Fact]
    public void Blank_option_values_are_dropped()
    {
        var argv = OneMoreCommand.GetHierarchy(notebook: "   ", section: "S").Build();
        Assert.Equal(new[] { "GetHierarchy", "--section", "S" }, argv);
    }
}
