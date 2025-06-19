using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:any>]]></c> inside <c><![CDATA[<xsd:complexType>]]></c> (through
/// <c><![CDATA[<xsd:sequence>]]></c> or <c><![CDATA[<xsd:choice>]]></c>).
/// <example>
/// <code><![CDATA[<xsd:any processContents="lax"/>]]></code>
/// </example>
/// </summary>
public class Any : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    public required ProcessContents ProcessContent { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
