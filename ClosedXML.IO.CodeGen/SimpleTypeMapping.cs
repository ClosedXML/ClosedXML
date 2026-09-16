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
    /// C# code template for calling a method that gets the value of a required attribute. The template method call is
    /// appended by <c>("attributeName")</c> for unqualified attribute or <c>("attributeName", "namespace")</c>
    /// for qualified attribute.
    /// </summary>
    public string RequiredTemplate
    {
        get => field ?? throw new InvalidOperationException($"Required template not defined for {Name}.");
        init;
    }

    /// <summary>
    /// C# code template for calling a method that gets the value of an optional  attribute. The template method call
    /// is appended by <c>("attributeName")</c> for unqualified attribute or <c>("attributeName", "namespace")</c>
    /// for qualified attribute.
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
