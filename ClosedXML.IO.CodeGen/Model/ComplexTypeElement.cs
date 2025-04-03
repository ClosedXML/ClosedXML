using ClosedXML.IO.CodeGen.Model.TopLevel;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// An element whose content will be a <see cref="ComplexTypeSequence"/>.
/// <code><![CDATA[<xsd:element name="field" maxOccurs="unbounded" type="CT_Field"/>]]></code>
/// </summary>
public class ComplexTypeElement : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// Name of the element in XML.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// A reference to a <see cref="ComplexTypeSequence"/>.
    /// </summary>
    public required string TypeName { get; init; }

    public required Occurrences Occurrences { get; init; }
}
