using System.Xml.Linq;

namespace OneMoreMcp.Core;

/// <summary>
/// Well-known names from the OneNote 2013 page schema that OneMore's CLI reads and writes. The
/// <c>one:</c> prefix maps to <see cref="Ns"/> in every page/hierarchy document the CLI emits.
/// </summary>
public static class OneNoteSchema
{
    /// <summary>The OneNote 2013 XML namespace (prefix <c>one</c>).</summary>
    public static readonly XNamespace Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    public static XName Page => Ns + "Page";
    public static XName Title => Ns + "Title";
    public static XName Outline => Ns + "Outline";
    public static XName OEChildren => Ns + "OEChildren";
    public static XName OE => Ns + "OE";
    public static XName T => Ns + "T";
    public static XName Position => Ns + "Position";
    public static XName Size => Ns + "Size";
}
