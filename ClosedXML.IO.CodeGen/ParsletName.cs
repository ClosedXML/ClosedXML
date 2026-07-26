using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A name for a top-level elements: <see cref="ComplexType"/>, <see cref="GroupDefinition"/> and <see cref="AttributeGroupDefinition"/>.
/// </summary>
public readonly record struct ParsletName
{
    public const string CtPrefix = "CT_";
    public const string EgPrefix = "EG_";
    public const string AgPrefix = "AG_";

    private static readonly HashSet<string> Special = ["xsd:string"];

    private ParsletName(string name)
    {
        if (!IsValidName(name))
            throw new ArgumentException($"Name '{name}' doesn't fit pattern for complex type or element group.");

        Value = name;
    }

    internal string Value { get; }

    /// <summary>
    /// Does it include a namespace. That generally means it's a reference to another XSD.
    /// </summary>
    internal bool HasNamespace => Value.Length > 3 && Value[1] == ':';

    public static implicit operator ParsletName(string name) => new(name);

    public string WithoutPrefix()
    {
        return Value[3..];
    }

    public override string ToString()
    {
        return Value;
    }

    private static bool IsValidName(string name)
    {
        if (Special.Contains(name))
            return true;

        if (name.Length > 3 && name[1] == ':')
            name = name[2..];

        return name.StartsWith(CtPrefix) ||
               name.StartsWith(EgPrefix) ||
               name.StartsWith(AgPrefix);
    }
}
