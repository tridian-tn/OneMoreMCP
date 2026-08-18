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
    public void AddHashtag_includes_the_required_notebook()
    {
        var argv = OneMoreCommand.AddHashtag("#todo", notebook: "Work", section: "S").Build();
        Assert.Equal(new[] { "AddHashtag", "--notebook", "Work", "--section", "S", "--tags", "#todo" }, argv);
    }

    [Fact]
    public void Search_includes_required_notebook_and_optional_scope()
    {
        var argv = OneMoreCommand.Search("meeting", notebook: "Work", section: "Ideas").Build();
        Assert.Equal(new[] { "Search", "--notebook", "Work", "--section", "Ideas", "--query", "meeting" }, argv);
    }

    [Fact]
    public void SearchTitles_includes_query_and_optional_notebook()
    {
        Assert.Equal(
            new[] { "SearchTitles", "--query", "agenda", "--notebook", "Work" },
            OneMoreCommand.SearchTitles("agenda", notebook: "Work").Build());
        Assert.Equal(new[] { "SearchTitles", "--query", "agenda" }, OneMoreCommand.SearchTitles("agenda").Build());
    }

    [Fact]
    public void Sync_targets_the_notebook()
    {
        Assert.Equal(new[] { "Sync", "--notebook", "Work" }, OneMoreCommand.Sync("Work").Build());
    }

    [Fact]
    public void Archive_carries_notebook_outfile_and_optional_section()
    {
        Assert.Equal(
            new[] { "Archive", "--notebook", "Work", "--section", "Ideas", "--outfile", @"C:\b\w.zip" },
            OneMoreCommand.Archive("Work", "Ideas", @"C:\b\w.zip").Build());
        Assert.Equal(
            new[] { "Archive", "--notebook", "Work", "--outfile", @"C:\b\w.zip" },
            OneMoreCommand.Archive("Work", null, @"C:\b\w.zip").Build());
    }

    [Fact]
    public void Diagnostics_adds_windows_switch_only_when_requested()
    {
        Assert.Equal(new[] { "Diagnostics" }, OneMoreCommand.Diagnostics().Build());
        Assert.Equal(new[] { "Diagnostics", "--windows" }, OneMoreCommand.Diagnostics(includeWindows: true).Build());
    }

    [Fact]
    public void Cleanup_scopes_to_notebook_section_page()
    {
        Assert.Equal(
            new[] { "RemoveEmpty", "--notebook", "Work", "--section", "Ideas" },
            OneMoreCommand.Cleanup("RemoveEmpty", "Work", "Ideas").Build());
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
    public void Export_backup_adds_the_bare_switch()
    {
        var argv = OneMoreCommand.Export(outpath: @"C:\out", format: "PDF", backup: true).Build();
        Assert.Equal(new[] { "Export", "--outpath", @"C:\out", "--format", "PDF", "--backup" }, argv);
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
