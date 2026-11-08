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
