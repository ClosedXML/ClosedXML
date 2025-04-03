using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// A representation of a <c><![CDATA[<xsd:choice>]]></c> element.
/// </summary>
public class Choice : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }
}
