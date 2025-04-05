using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

public abstract class ComplexType : IReferencable
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required string Name { get; set; }

    public List<OneOf<AttributeElement, AttributeGroupReference>> Attributes { get; set; } = [];
}
