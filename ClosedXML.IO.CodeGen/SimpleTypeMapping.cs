using System;

namespace ClosedXML.IO.CodeGen;

internal record SimpleTypeMapping
{
    private readonly string? _requiredTemplate;
    private readonly string? _optionalTemplate;

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
        get => _requiredTemplate ?? throw new InvalidOperationException();
        init => _requiredTemplate = value;
    }

    /// <summary>
    /// C# code template for getting a value from an optional attribute. The name of attribute is in the string as <c>{0}</c>.
    /// </summary>
    public string OptionalTemplate
    {
        get => _optionalTemplate ?? throw new InvalidOperationException();
        init => _optionalTemplate = value;
    }

    /// <summary>
    /// Map values from XML default value to C# value.
    /// </summary>
    public Func<string, string> MapValue { get; init; } = x => x;
}
