using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A name for a elements with a type that can be converted to <c>Parse*</c> calls.
/// </summary>
public readonly record struct ParsletName
{
    /// <summary>Element has type of <see cref="ComplexType"/></summary>
    private const string CtPrefix = "CT_";

    /// <summary>Element has type of <see cref="GroupDefinition"/></summary>
    private const string EgPrefix = "EG_";

    /// <summary>Element has type of <see cref="AttributeGroupDefinition"/></summary>
    private const string AgPrefix = "AG_";

    /// <summary>Element with a content that should be converted to some type (mostly string).</summary>
    private const string StPrefix = "ST_";

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
               name.StartsWith(AgPrefix) ||
               name.StartsWith(StPrefix);
    }
}
