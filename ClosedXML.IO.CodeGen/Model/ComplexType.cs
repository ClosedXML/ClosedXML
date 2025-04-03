using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A complex type that uses <see cref="SequenceElement"/>/<see cref="ChoiceElement"/> and
/// attributes to define the complex type.
/// </summary>
public class ComplexType : ComplexTypeBase
{
    public required List<IElementGroup> ElementGroups { get; init; } = [];
}
