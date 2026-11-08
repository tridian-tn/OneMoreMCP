using System.Globalization;
using System.Xml.Linq;
using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>
/// Covers the append engine: existing content is preserved and exactly the new paragraphs are added,
/// across the page-with-outline, page-without-outline, tag-bank, target-choice, and error cases.
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

    private static string Placed(string text, string? y, string? height = null, string? objectId = null, string z = "0") =>
        "<one:Outline" + (objectId is null ? "" : $" objectID=\"{objectId}\"") + ">" +
        (y is null ? "" : $"<one:Position x=\"144.0\" y=\"{y}\" z=\"{z}\" />") +
        (height is null ? "" : $"<one:Size width=\"72.0\" height=\"{height}\" />") +
        $"<one:OEChildren><one:OE><one:T><![CDATA[{text}]]></one:T></one:OE></one:OEChildren></one:Outline>";

    private static string AsTagBank(string outline) =>
        outline.Replace("<one:OEChildren>", "<one:Meta name=\"omTaggingBank\" content=\"1\" /><one:OEChildren>");

    private static IEnumerable<string[]> Orderings(string[] outlines) =>
        outlines.Length <= 1
            ? new[] { outlines }
            : outlines.SelectMany(first => Orderings(outlines.Where(o => o != first).ToArray())
                .Select(rest => new[] { first }.Concat(rest).ToArray()));

    /// <summary>Checks the same outline takes the text in every XML order, naming it by its first run.</summary>
    /// <param name="expected">The first run of the outline that should take the text</param>
    /// <param name="outlines">The page's outlines, each with a different first run</param>
    private static void AssertTarget(string expected, params string[] outlines)
    {
        foreach (var ordering in Orderings(outlines))
        {
            var result = PageAppender.Append(PageWith(ordering), "New line", AppendFormat.Plain);
            Assert.Equal(expected, RunsByOutline(result).Single(runs => runs.Contains("New line"))[0]);
        }
    }

    /// <summary>Does to the XML what OneNote does after a write: outlines swap in both XML order and z-order.</summary>
    /// <param name="xml">The page as it was written</param>
    /// <returns>The page as OneNote would hand it back</returns>
    private static string Reshuffled(string xml)
    {
        var doc = XDocument.Parse(xml);
        var outlines = doc.Root!.Elements(One + "Outline").ToList();
        var zs = outlines.Select(o => o.Element(One + "Position")!.Attribute("z")!.Value).Reverse().ToList();
        for (var i = 0; i < outlines.Count; i++) outlines[i].Element(One + "Position")!.SetAttributeValue("z", zs[i]);

        outlines.Remove();
        doc.Root.Add(Enumerable.Reverse(outlines));
        return doc.ToString();
    }

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
    public void Append_goes_into_the_lowest_outline_whatever_the_xml_order()
    {
        AssertTarget(
            "lowest",
            Placed("top", "70.0", "50.0"),
            Placed("lowest", "400.0", "20.0"),
            Placed("middle", "200.0", "50.0"));
    }

    [Theory]
    [InlineData("300.0", "30.0", "main")] // a side note starting lower down, but ending well above the main note
    [InlineData("900.0", "20.0", "other")] // a note under the main one
    public void Outlines_are_compared_by_where_they_end(string otherY, string otherHeight, string expected)
    {
        AssertTarget(expected, Placed("main", "70.0", "800.0"), Placed("other", otherY, otherHeight));
    }

    [Fact]
    public void Outlines_in_the_same_spot_go_to_the_taller_one()
    {
        AssertTarget(
            "taller",
            Placed("shorter", "266.4", "40.28", "{A}{14}{B0}"),
            Placed("taller", "266.4", "93.99", "{A}{10}{B0}"));
    }

    [Fact]
    public void Outlines_in_the_same_spot_at_the_same_size_go_by_object_id()
    {
        AssertTarget(
            "second",
            Placed("first", "266.4", "93.99", "{A}{10}{B0}"),
            Placed("second", "266.4", "93.99", "{A}{14}{B0}"));
    }

    [Fact]
    public void Repeated_appends_stay_in_one_outline_when_onenote_reorders_the_page()
    {
        var xml = PageWith(
            Placed("first", "266.4", "93.99", "{A}{10}{B0}", z: "0"),
            Placed("second", "266.4", "93.99", "{A}{14}{B0}", z: "1"));

        foreach (var line in new[] { "one", "two", "three", "four" })
        {
            xml = Reshuffled(PageAppender.Append(xml, line, AppendFormat.Plain));
        }

        var outlines = RunsByOutline(xml);
        Assert.Equal(new[] { "first" }, outlines.Single(runs => runs[0] == "first"));
        Assert.Equal(new[] { "second", "one", "two", "three", "four" }, outlines.Single(runs => runs[0] == "second"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("not a number")]
    public void An_outline_with_no_usable_position_loses_to_one_that_has_it(string? badY)
    {
        AssertTarget("placed", Placed("unplaced", badY, "900.0"), Placed("placed", "10.0", "1.0"));
    }

    [Fact]
    public void Outlines_without_positions_go_by_object_id_if_they_have_one()
    {
        AssertTarget(
            "second",
            Placed("first", null, objectId: "{A}{10}{B0}"),
            Placed("second", null, objectId: "{A}{14}{B0}"));
    }

    [Fact]
    public void Outlines_with_nothing_to_compare_fall_back_to_the_last_in_the_xml()
    {
        var result = PageAppender.Append(PageWith(Placed("first", null), Placed("second", null)), "New line", AppendFormat.Plain);

        Assert.Equal(new[] { "second", "New line" }, RunsByOutline(result)[1]);
    }

    [Theory]
    [InlineData("200.0", "first")] // first ends at 300, second at 250
    [InlineData("300.0", "second")] // first ends at 300, second at 350
    public void An_outline_with_no_size_ends_where_it_starts(string secondHeight, string expected)
    {
        AssertTarget(expected, Placed("first", "300.0"), Placed("second", "50.0", secondHeight));
    }

    [Fact]
    public void Positions_are_read_the_same_under_any_culture()
    {
        // In German '.' isn't a decimal point, so a culture-sensitive parse would misread or reject these.
        var original = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = new CultureInfo("de-DE");
        try
        {
            AssertTarget("lower", Placed("lower", "266.5", "1.5"), Placed("upper", "100.5", "10.5"));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void Append_skips_a_tag_bank_that_ends_lowest()
    {
        AssertTarget("body", Placed("body", "100.0", "50.0"), AsTagBank(Placed("#todo", "900.0", "50.0")));
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
