using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>Asserts the exact argument vectors the command factories produce.</summary>
public class OneMoreCommandTests
{
    [Fact]
    public void GetHierarchy_includes_only_supplied_options()
    {
        var argv = OneMoreCommand.GetHierarchy(notebook: "Work", section: null, books: true).Build();
        Assert.Equal(new[] { "GetHierarchy", "--notebook", "Work", "--books", "yes" }, argv);
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
        Assert.Equal(new[] { "GetPage", "--current", "yes" }, argv);
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
        Assert.Equal(new[] { "PutPage", "--notebook", "B", "--section", "S", "--infile", @"C:\t\p.xml", "--force", "yes" }, argv);
    }

    [Fact]
    public void AddHashtag_includes_the_required_notebook()
    {
        var argv = OneMoreCommand.AddHashtag("#todo", notebook: "Work", section: "S").Build();
        Assert.Equal(new[] { "AddHashtag", "--notebook", "Work", "--section", "S", "--tags", "#todo" }, argv);
    }

    [Fact]
    public void SearchHashtags_allTags_adds_the_switch()
    {
        Assert.Equal(
            new[] { "SearchHashtags", "--query", "#a #b", "--allTags", "yes" },
            OneMoreCommand.SearchHashtags("#a #b", allTags: true).Build());
    }

    [Fact]
    public void InsertToc_always_emits_refresh_yes_or_no()
    {
        Assert.Equal(
            new[] { "InsertToc", "--notebook", "B", "--section", "S", "--page", "P", "--refresh", "no" },
            OneMoreCommand.InsertToc("B", "S", "P", refresh: false).Build());
        Assert.Equal(
            new[] { "InsertToc", "--notebook", "B", "--section", "S", "--page", "P", "--refresh", "yes" },
            OneMoreCommand.InsertToc("B", "S", "P", refresh: true).Build());
    }

    [Fact]
    public void Export_includes_format_and_optional_pageId()
    {
        var argv = OneMoreCommand.Export(outpath: @"C:\out", format: "PDF", pageId: "{123}").Build();
        Assert.Equal(new[] { "Export", "--outpath", @"C:\out", "--format", "PDF", "--pageId", "{123}" }, argv);
    }

    [Fact]
    public void Output_appends_the_global_output_option()
    {
        var argv = OneMoreCommand.GetPage(notebook: "N", section: "S", page: "P", current: false)
            .Output(@"C:\tmp\o.xml").Build();
        Assert.Equal(
            new[] { "GetPage", "--notebook", "N", "--section", "S", "--page", "P", "--output", @"C:\tmp\o.xml" },
            argv);
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
