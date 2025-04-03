using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A complex type that uses <see cref="SequenceElement"/> as a root element and
/// attributes to define the complex type.
/// </summary>
public class ComplexTypeSequence : ComplexType
{
    public required List<IElementGroup> Elements { get; init; } = [];
}
