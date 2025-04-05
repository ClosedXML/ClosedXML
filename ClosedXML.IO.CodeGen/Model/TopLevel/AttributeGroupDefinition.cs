using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:attributeGroup/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
///  <xsd:attributeGroup name="AG_RevData">
///    <xsd:attribute name="rId" type="xsd:unsignedInt" use="required"/>
///    <xsd:attribute name = "ua" type="xsd:boolean" use="optional" default="false"/>
///    <xsd:attribute name = "ra" type="xsd:boolean" use="optional" default="false"/>
///  </xsd:attributeGroup>
/// ]]></code>
/// </example>
/// </summary>
public class AttributeGroupDefinition : IReferencable, INode
{
    /// <summary>
    /// Name of the the attribute group type.
    /// </summary>
    public required string Name { get; init; }

    public required List<AttributeElement> Attributes { get; init; } = [];

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
