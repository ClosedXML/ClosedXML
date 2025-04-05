using ClosedXML.IO.CodeGen.Model.TopLevel;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:element ref="some:element">]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// <example>
/// <code><![CDATA[
///   <xsd:element ref="xdr:from" minOccurs="1" maxOccurs="1"/>
/// ]]></code>
/// </example>
/// </summary>
public class ElementReference : ILeafElement
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// Name of referenced element in the element definition (<see cref="ElementDefinition.Name"/>).
    /// </summary>
    public required string RefName { get; init; }

    public required Occurrences Occurrences { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
