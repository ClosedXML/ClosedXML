using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// A definition of a restriction of a <see cref="ISimpleType"/> through a base type and additional
/// restrictions of possible values.
/// </summary>
public class Restriction
{
    /// <summary>
    /// Name of a base type, e.g., <c>xsd:string</c>.
    /// </summary>
    public required string BaseTypeName {get; init; }

    /// <summary>
    /// Additional restrictions on possible values.
    /// </summary>
    public required List<IValueRestriction> ValueRestrictions { get; init; }
}
