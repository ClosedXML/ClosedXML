using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:group ref="some:group">]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class GroupReference : ILeafElement
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// A reference to the element (<see cref="GroupDefinition.Name"/>).
    /// </summary>
    public required string RefName { get; init; }

    public required Occurrences Occurrences { get; init; }
}
