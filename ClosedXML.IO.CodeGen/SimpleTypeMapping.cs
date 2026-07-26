using System;

namespace ClosedXML.IO.CodeGen;

public record SimpleTypeMapping
{
    /// <summary>
    /// Name of the simple type in the XML.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// Name of the mapped C# type.
    /// </summary>
    public required string CsTypeName { get; init; }

    /// <summary>
    /// C# code template for getting a value from a required attribute. The name of attribute is in the string as <c>{0}</c>.
    /// </summary>
    public string RequiredTemplate
    {
        get => field ?? throw new InvalidOperationException($"Required template not defined for {Name}.");
        init;
    }

    /// <summary>
    /// C# code template for getting a value from an optional attribute. The name of attribute is in the string as <c>{0}</c>.
    /// </summary>
    public string OptionalTemplate
    {
        get => field ?? throw new InvalidOperationException($"Optional template not defined for {Name}.");
        init;
    }

    /// <summary>
    /// Map values from XML default value to C# value.
    /// </summary>
    public Func<string, string> MapValue { get; init; } = x => x;
}
