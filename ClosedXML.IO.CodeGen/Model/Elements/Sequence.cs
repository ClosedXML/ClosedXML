using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// A representation of a <c><![CDATA[<xsd:sequence>]]></c> element.
/// </summary>
public class Sequence : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }
}
