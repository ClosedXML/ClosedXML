using ClosedXML.IO.CodeGen.Model.TopLevel;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// A reference to an element in the root of <see cref="Schema"/>.
/// <code><![CDATA[<xsd:element ref="xdr:from" minOccurs="1" maxOccurs="1"/>]]></code>
/// </summary>
public class ElementReference : ILeafElement
{
    /// <summary>
    /// Name of referenced element in the element definition (<see cref="ElementDefinition.Name"/>).
    /// </summary>
    public required string RefName { get; init; }

    public required Occurrences Occurrences { get; init; }

    public List<IElementGroup> Children { get; } = [];
}
