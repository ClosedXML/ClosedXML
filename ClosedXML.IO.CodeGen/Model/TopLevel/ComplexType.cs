using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>. It doesn't have
/// any elements, only attributes.
/// </summary>
public class ComplexType : IReferencable
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required string Name { get; set; }

    public List<OneOf<AttributeElement, AttributeGroupReference>> Attributes { get; set; } = [];
}
