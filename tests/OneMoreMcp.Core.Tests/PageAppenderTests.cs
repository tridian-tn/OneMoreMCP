using System.Xml.Linq;
using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>
/// Covers the append engine: existing content is preserved and exactly the new paragraphs are added,
/// across the page-with-outline, page-without-outline, tag-bank, and error cases.
/// </summary>
public class PageAppenderTests
{
    private static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private const string PageWithOutline =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\">" +
        "<one:Title><one:OE><one:T><![CDATA[My Title]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[Existing line]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline>" +
        "</one:Page>";

    private const string PageNoOutline =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P2}\">" +
        "<one:Title><one:OE><one:T><![CDATA[Bare]]></one:T></one:OE></one:Title>" +
        "</one:Page>";

    private const string BodyOutline =
        "<one:Outline><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[Existing line]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline>";

    private static string TagBankOutline(string markerValue = "1") =>
        $"<one:Outline><one:Meta name=\"omTaggingBank\" content=\"{markerValue}\" /><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[#todo]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline>";

    private static string PageWith(params string[] outlines) =>
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P3}\">" +
        "<one:Title><one:OE><one:T><![CDATA[Tagged]]></one:T></one:OE></one:Title>" +
        string.Concat(outlines) +
        "</one:Page>";

    private static List<string> Runs(string xml) =>
        XDocument.Parse(xml)
            .Descendants(One + "Outline")
            .Descendants(One + "T")
            .Select(t => t.Value)
            .ToList();

    private static List<List<string>> RunsByOutline(string xml) =>
        XDocument.Parse(xml)
            .Descendants(One + "Outline")
            .Select(outline => outline.Descendants(One + "T").Select(t => t.Value).ToList())
            .ToList();

    [Fact]
    public void Append_keeps_existing_content_and_adds_the_new_paragraph()
    {
        var result = PageAppender.Append(PageWithOutline, "New line", AppendFormat.Plain);
        var runs = Runs(result);
        Assert.Contains("Existing line", runs);
        Assert.Contains("New line", runs);
        Assert.Equal(2, runs.Count); // one existing + one appended, nothing dropped
    }

    [Fact]
    public void Append_preserves_the_page_id()
    {
        var result = PageAppender.Append(PageWithOutline, "x", AppendFormat.Plain);
        Assert.Equal("{P1}", (string?)XDocument.Parse(result).Root!.Attribute("ID"));
    }

    [Fact]
    public void Append_splits_multiple_lines_into_multiple_paragraphs()
    {
        var result = PageAppender.Append(PageWithOutline, "a\nb\nc", AppendFormat.Plain);
        Assert.Equal(new[] { "Existing line", "a", "b", "c" }, Runs(result));
    }

    [Fact]
    public void Plain_text_is_html_encoded()
    {
        var result = PageAppender.Append(PageWithOutline, "a < b & c", AppendFormat.Plain);
        Assert.Contains("a &lt; b &amp; c", Runs(result));
    }

    [Fact]
    public void Markdown_bold_becomes_a_bold_span()
    {
        var result = PageAppender.Append(PageWithOutline, "**hi** there", AppendFormat.Markdown);
        Assert.Contains(Runs(result), r => r.Contains("font-weight:bold") && r.Contains("hi"));
    }

    [Fact]
    public void Html_is_placed_verbatim()
    {
        var result = PageAppender.Append(PageWithOutline, "<b>raw</b>", AppendFormat.Html);
        Assert.Contains("<b>raw</b>", Runs(result));
    }

    [Fact]
    public void Append_to_a_page_without_an_outline_is_refused()
    {
        // OneNote ignores a new top-level outline, so building one would only produce a write it discards.
        var ex = Assert.Throws<ArgumentException>(
            () => PageAppender.Append(PageNoOutline, "first content", AppendFormat.Plain));

        Assert.Contains("no outline", ex.Message);
        Assert.Contains("typing a line", ex.Message);
        Assert.DoesNotContain("AddHashtag", ex.Message); // text in a tag bank isn't safe, so don't send people there
        Assert.DoesNotContain("Parameter", ex.Message); // this message reaches the client as-is
    }

    [Fact]
    public void Append_goes_into_the_body_when_the_tag_bank_comes_first()
    {
        // This is where OneMore puts a new bank: straight after the title.
        var result = PageAppender.Append(PageWith(TagBankOutline(), BodyOutline), "New line", AppendFormat.Plain);

        var outlines = RunsByOutline(result);
        Assert.Equal(new[] { "#todo" }, outlines[0]);
        Assert.Equal(new[] { "Existing line", "New line" }, outlines[1]);
    }

    [Fact]
    public void Append_skips_a_tag_bank_that_comes_last()
    {
        var result = PageAppender.Append(PageWith(BodyOutline, TagBankOutline()), "New line", AppendFormat.Plain);

        var outlines = RunsByOutline(result);
        Assert.Equal(new[] { "Existing line", "New line" }, outlines[0]);
        Assert.Equal(new[] { "#todo" }, outlines[1]);
    }

    [Theory]
    [InlineData("1")]
    [InlineData("")] // OneMore only checks the marker's name, so a blank value is still a bank
    public void Append_to_a_page_whose_only_outline_is_a_tag_bank_is_refused(string markerValue)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PageAppender.Append(PageWith(TagBankOutline(markerValue)), "note", AppendFormat.Plain));

        Assert.Contains("tag bank", ex.Message);
        Assert.Contains("typing a line", ex.Message);
        Assert.DoesNotContain("Parameter", ex.Message); // this message reaches the client as-is
    }

    [Fact]
    public void An_outline_with_some_other_meta_still_takes_text()
    {
        var outline = TagBankOutline().Replace("omTaggingBank", "omSomethingElse");

        var result = PageAppender.Append(PageWith(outline), "New line", AppendFormat.Plain);

        Assert.Equal(new[] { "#todo", "New line" }, Runs(result));
    }

    [Fact]
    public void Non_page_xml_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => PageAppender.Append("<foo/>", "x", AppendFormat.Plain));
    }

    [Fact]
    public void Unparseable_xml_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => PageAppender.Append("<one:Page", "x", AppendFormat.Plain));
    }

    [Fact]
    public void Empty_text_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => PageAppender.Append(PageWithOutline, "", AppendFormat.Plain));
    }
}
