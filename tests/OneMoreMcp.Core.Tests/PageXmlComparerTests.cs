using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>
/// Covers the comparison behind the post-write check: attributes that move without the content changing
/// are ignored, and any real change is still seen.
/// </summary>
public class PageXmlComparerTests
{
    private const string SamplePage =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\" name=\"P\" " +
        "lastModifiedTime=\"2026-11-08T15:31:36.000Z\" isCurrentlyViewed=\"true\" selected=\"partial\">" +
        "<one:Title selected=\"partial\" omHash=\"111\">" +
        "<one:OE lastModifiedTime=\"2026-11-08T15:00:00.000Z\" quickStyleIndex=\"0\" selected=\"partial\">" +
        "<one:T selected=\"all\"><![CDATA[My Title]]></one:T>" +
        "</one:OE></one:Title>" +
        "<one:Outline omHash=\"222\"><one:OEChildren>" +
        "<one:OE quickStyleIndex=\"1\"><one:T><![CDATA[Existing line]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline></one:Page>";

    /// <summary>Applies an edit to the sample page, failing if the edit didn't match anything.</summary>
    /// <param name="find">The text to replace</param>
    /// <param name="replacement">The text to put in its place</param>
    /// <returns>The edited page XML</returns>
    private static string Edited(string find, string replacement)
    {
        Assert.Contains(find, SamplePage);
        return SamplePage.Replace(find, replacement);
    }

    [Theory]
    [InlineData("lastModifiedTime=\"2026-11-08T15:31:36.000Z\"", "lastModifiedTime=\"2026-11-08T15:32:10.000Z\"")] // on the page
    [InlineData("lastModifiedTime=\"2026-11-08T15:00:00.000Z\"", "lastModifiedTime=\"2026-11-08T15:32:10.000Z\"")] // on a paragraph
    [InlineData(" isCurrentlyViewed=\"true\"", "")]
    [InlineData("<one:Title selected=\"partial\"", "<one:Title")]
    [InlineData("<one:T selected=\"all\">", "<one:T>")]
    [InlineData("omHash=\"111\"", "omHash=\"999\"")]
    public void Volatile_attributes_are_ignored(string find, string replacement)
    {
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, Edited(find, replacement)));
    }

    [Theory]
    [InlineData("Existing line", "Edited line")]
    [InlineData("My Title", "New Title")]
    [InlineData("Existing line]]></one:T></one:OE>", "Existing line]]></one:T></one:OE><one:OE><one:T><![CDATA[Added]]></one:T></one:OE>")]
    [InlineData("quickStyleIndex=\"1\"", "quickStyleIndex=\"2\"")]
    [InlineData("name=\"P\"", "name=\"Q\"")]
    [InlineData("Existing line", "Existing line ")] // whitespace inside a text run is content
    [InlineData("Existing line", "<!-- note -->")] // and so is a comment in the run's HTML
    [InlineData("<one:Outline ", "<one:Outline ID=\"{X}\" ")] // only the root's ID is ignored
    public void Real_changes_are_seen(string find, string replacement)
    {
        Assert.False(PageXmlComparer.HasSameContent(SamplePage, Edited(find, replacement)));
    }

    [Fact]
    public void Formatting_whitespace_is_ignored()
    {
        var indented = Edited("<one:OEChildren>", "\n    <one:OEChildren>\n      ");
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, indented));
    }

    [Theory]
    [InlineData("<![CDATA[Existing line]]>", "Existing line")]
    [InlineData("<![CDATA[Existing line]]>", "<![CDATA[Existing]]> line")] // split across nodes
    public void Cdata_and_plain_text_with_the_same_characters_match(string find, string replacement)
    {
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, Edited(find, replacement)));
    }

    [Fact]
    public void An_empty_cdata_section_matches_an_empty_element()
    {
        var empty = Edited("<![CDATA[Existing line]]>", "<![CDATA[]]>");
        Assert.True(PageXmlComparer.HasSameContent(empty, empty.Replace("<![CDATA[]]>", "")));
    }

    [Theory]
    [InlineData("<one:Page ", "<!-- note --><one:Page ")]
    [InlineData("<one:Title", "<!-- note --><one:Title")]
    [InlineData("</one:Page>", "</one:Page><!-- note -->")]
    [InlineData("<one:Page ", "<?note?><one:Page ")]
    [InlineData("<one:Title", "<?note?><one:Title")]
    [InlineData("<one:Page ", "<!DOCTYPE note><one:Page ")]
    public void Nodes_that_never_reach_the_page_are_ignored(string find, string replacement)
    {
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, Edited(find, replacement)));
    }

    [Theory]
    [InlineData("ID=\"{P1}\"", "ID=\"{Q1}\"")]
    [InlineData(" ID=\"{P1}\"", "")]
    public void The_root_id_is_ignored(string find, string replacement)
    {
        // PutPage replaces it with the ID of the page it's writing to, so it never reaches the page as sent.
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, Edited(find, replacement)));
    }

    [Fact]
    public void Attribute_order_is_ignored()
    {
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, Edited("ID=\"{P1}\" name=\"P\"", "name=\"P\" ID=\"{P1}\"")));
    }

    [Fact]
    public void A_whitespace_only_plain_text_run_matches_an_empty_run()
    {
        // OneMore parses the XML it's given with the same defaults and drops this space, so it writes an empty run.
        var empty = Edited("<![CDATA[Existing line]]>", "");
        Assert.True(PageXmlComparer.HasSameContent(empty, Edited("<![CDATA[Existing line]]>", " ")));
    }

    [Fact]
    public void A_whitespace_only_plain_text_run_differs_from_a_cdata_space()
    {
        // OneNote keeps a space held in CDATA, and writing it back as plain text really does remove it.
        var cdataSpace = Edited("<![CDATA[Existing line]]>", "<![CDATA[ ]]>");
        Assert.False(PageXmlComparer.HasSameContent(cdataSpace, Edited("<![CDATA[Existing line]]>", " ")));
    }

    [Fact]
    public void Text_is_compared_after_decoding()
    {
        // In CDATA "&lt;" is four literal characters; as plain text it's one "<".
        var escaped = Edited("Existing line]]>", "a < b]]>");
        Assert.True(PageXmlComparer.HasSameContent(escaped, escaped.Replace("<![CDATA[a < b]]>", "a &lt; b")));
        Assert.False(PageXmlComparer.HasSameContent(escaped, escaped.Replace("<![CDATA[a < b]]>", "<![CDATA[a &lt; b]]>")));
    }

    [Fact]
    public void A_different_namespace_prefix_matches()
    {
        var prefixed = SamplePage.Replace("one:", "x:").Replace("xmlns:one=", "xmlns:x=");
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, prefixed));
    }

    [Fact]
    public void A_default_namespace_matches()
    {
        var unprefixed = SamplePage.Replace("one:", "").Replace("xmlns:one=", "xmlns=");
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, unprefixed));
    }

    [Fact]
    public void An_extra_namespace_declaration_matches()
    {
        var extra = Edited("xmlns:one=", "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:one=");
        Assert.True(PageXmlComparer.HasSameContent(SamplePage, extra));
    }

    [Fact]
    public void A_different_namespace_is_a_real_change()
    {
        // The declarations can be ignored because each name keeps its namespace, so this must still differ.
        Assert.False(PageXmlComparer.HasSameContent(SamplePage, Edited("2013/onenote", "2010/onenote")));
    }

    [Fact]
    public void Identical_text_matches_even_when_it_is_not_xml()
    {
        Assert.True(PageXmlComparer.HasSameContent("No page is currently active.", "No page is currently active."));
    }

    [Theory]
    [InlineData("No page is currently active.")]
    [InlineData("")]
    public void Text_that_is_not_xml_does_not_match_a_page(string notXml)
    {
        Assert.False(PageXmlComparer.HasSameContent(SamplePage, notXml));
        Assert.False(PageXmlComparer.HasSameContent(notXml, SamplePage));
    }
}
