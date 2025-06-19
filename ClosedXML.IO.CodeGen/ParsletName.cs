using System;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A name for a top-level elements: <see cref="ComplexType"/> and <see cref="GroupDefinition"/>.
/// </summary>
internal readonly record struct ParsletName
{
    public const string CtPrefix = "CT_";
    public const string EgPrefix = "EG_";

    private ParsletName(string name)
    {
        if (!name.StartsWith(CtPrefix) && !name.StartsWith(EgPrefix))
            throw new ArgumentException($"Name '{name}' doesn't fit pattern for complex type or element group.");

        Value = name;
    }

    internal string Value { get; }

    public static implicit operator ParsletName(string name) => new(name);
}
