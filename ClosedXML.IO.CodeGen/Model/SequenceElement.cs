using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A representation of a <c><![CDATA[<xsd:sequence>]]></c> element.
/// </summary>
public class SequenceElement : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }
}
