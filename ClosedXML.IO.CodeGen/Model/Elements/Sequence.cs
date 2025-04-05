using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:sequence>]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class Sequence : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
