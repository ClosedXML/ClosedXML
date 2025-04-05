using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

public class SimpleTypeUnion : ISimpleType
{
    public required string Name { get; init; }

    public required List<Restriction> RestrictionsUnion { get; init; } = [];

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
