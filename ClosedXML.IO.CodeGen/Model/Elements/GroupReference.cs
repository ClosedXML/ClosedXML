using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// A reference to a defined element group (<see cref="GroupDefinition"/>).
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
