using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:choice>]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class Choice : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }
}
