using System.Xml.Linq;
using OneMoreMcp.Core;

namespace OneMoreMcp.Core.Tests;

/// <summary>Covers the PutPage sanitiser that strips attributes OneNote's write schema rejects.</summary>
public class OneNotePageTests
{
    private const string PageWithOmHash =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\">" +
        "<one:PageSettings omHash=\"ABC123\"><one:PageSize omHash=\"DEF\" /></one:PageSettings>" +
        "<one:Outline omHash=\"GHI\"><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[Keep me]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline></one:Page>";

    [Fact]
    public void PrepareForPut_removes_every_omHash_attribute()
    {
        var result = OneNotePage.PrepareForPut(PageWithOmHash);
        Assert.DoesNotContain("omHash", result);
    }

    [Fact]
    public void PrepareForPut_preserves_content_and_page_id()
    {
        var result = OneNotePage.PrepareForPut(PageWithOmHash);
        var doc = XDocument.Parse(result);
        Assert.Equal("{P1}", (string?)doc.Root!.Attribute("ID"));
        Assert.Contains("Keep me", result);
    }

    [Fact]
    public void PrepareForPut_rejects_unparseable_xml()
    {
        Assert.Throws<ArgumentException>(() => OneNotePage.PrepareForPut("<one:Page"));
    }
}
