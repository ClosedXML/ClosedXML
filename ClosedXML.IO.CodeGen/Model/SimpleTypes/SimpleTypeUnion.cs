using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

public class SimpleTypeUnion : ISimpleType
{
    public required string Name { get; init; }

    public required List<Restriction> RestrictionsUnion { get; init; } = [];
}
