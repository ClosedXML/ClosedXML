using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A base class for <c><![CDATA[<xsd:complexType/>]]></c>. It doesn't have any elements.
/// </summary>
public class ComplexType : IReferencable
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required string Name { get; set; }

    public List<AttributeElement> Attributes { get; set; } = [];
}
