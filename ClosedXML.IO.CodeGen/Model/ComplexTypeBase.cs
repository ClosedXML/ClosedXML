using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A base class for various types of a <c><![CDATA[<xsd:complexType/>]]></c>.
/// </summary>
public class ComplexTypeBase
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required string Name { get; set; }

    public List<AttributeElement> Attributes { get; set; } = [];
}
