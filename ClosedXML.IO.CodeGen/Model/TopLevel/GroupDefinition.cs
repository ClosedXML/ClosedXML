using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A referencable group of elements.
/// </summary>
internal class GroupDefinition : IReferencable
{
    public required string Name { get; init; }

    public required IElementGroup Content { get; init; }
}
