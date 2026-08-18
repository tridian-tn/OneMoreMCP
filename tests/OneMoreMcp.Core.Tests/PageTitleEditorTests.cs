using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>
/// Covers the title edit used by the two-step page create: the CLI makes an untitled shell, then the
/// title is applied by writing the page back, so the edit must preserve the IDs OneNote assigned.
/// </summary>
public class PageTitleEditorTests
{
    // A page as the CLI emits it right after creating one: an empty title, no outline, and an
    // objectID on the title's OE that the write back has to keep.
    private const string NewPageXml =
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\" name=\"Untitled\">" +
        "<one:Title omHash=\"abc\"><one:OE objectID=\"{O1}\"><one:T><![CDATA[]]></one:T></one:OE></one:Title>" +
        "</one:Page>";

    [Fact]
    public void Set_title_fills_an_empty_title()
    {
        var result = PageTitleEditor.SetTitle(NewPageXml, "My Page");
        Assert.Equal("My Page", PageTitleEditor.GetTitle(result));
    }

    [Fact]
    public void Set_title_preserves_the_ids_onenote_assigned()
    {
        // An update only applies to objects OneNote can match, so losing these silently discards it.
        var result = PageTitleEditor.SetTitle(NewPageXml, "My Page");
        Assert.Contains("ID=\"{P1}\"", result);
        Assert.Contains("objectID=\"{O1}\"", result);
    }

    [Fact]
    public void Set_title_replaces_existing_text_and_collapses_split_runs()
    {
        const string split =
            "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">" +
            "<one:Title><one:OE><one:T><![CDATA[Old ]]></one:T><one:T><![CDATA[title]]></one:T></one:OE></one:Title>" +
            "</one:Page>";

        var result = PageTitleEditor.SetTitle(split, "New");

        Assert.Equal("New", PageTitleEditor.GetTitle(result));
        Assert.DoesNotContain("Old", result);
    }

    [Fact]
    public void Set_title_adds_a_title_when_the_page_has_none()
    {
        const string bare = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" />";
        var result = PageTitleEditor.SetTitle(bare, "Fresh");
        Assert.Equal("Fresh", PageTitleEditor.GetTitle(result));
    }

    [Fact]
    public void Set_title_rejects_bad_input()
    {
        Assert.Throws<ArgumentException>(() => PageTitleEditor.SetTitle("", "T"));
        Assert.Throws<ArgumentException>(() => PageTitleEditor.SetTitle(NewPageXml, "  "));
        Assert.Throws<ArgumentException>(() => PageTitleEditor.SetTitle("<one:Page", "T"));
        Assert.Throws<ArgumentException>(() => PageTitleEditor.SetTitle("<Other/>", "T"));
    }

    [Fact]
    public void Get_title_returns_null_when_absent_or_empty()
    {
        Assert.Null(PageTitleEditor.GetTitle(NewPageXml));
        Assert.Null(PageTitleEditor.GetTitle("<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" />"));
        Assert.Null(PageTitleEditor.GetTitle("not xml"));
    }

    [Fact]
    public void New_page_xml_is_a_titled_one_page_document()
    {
        var xml = PageTitleEditor.NewPageXml("Seed");
        Assert.Contains("one:Page", xml);
        Assert.Equal("Seed", PageTitleEditor.GetTitle(xml));
    }
}
