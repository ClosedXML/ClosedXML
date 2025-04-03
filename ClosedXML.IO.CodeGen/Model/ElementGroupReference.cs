using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

public class ElementGroupReference : ILeafElement
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// A reference to the element.
    /// </summary>
    public required string RefName { get; init; }

    public required Occurrences Occurrences { get; init; }
}
