using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

internal class SimpleTypeEnum
{
    public required string Name { get; init; }

    public required string BaseTypeName { get; init; }

    public required List<string> Values { get; init; }

    public required int? Length { get; init; }

    public required int? MinInclusive { get; init; }

    public required int? MaxInclusive { get; init; }
}
