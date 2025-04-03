using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A complex type that uses <see cref="ChoiceElement"/> as a root element and
/// attributes to define the complex type.
/// </summary>
public class ComplexTypeChoice : ComplexType
{
    public required List<IElementGroup> Choices { get; init; } = [];
}
