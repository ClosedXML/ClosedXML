using System;
using ClosedXML.IO.CodeGen.Model.TopLevel;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:element ref="some:element">]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>
/// (either <c><![CDATA[<xsd:sequence>]]></c> or <c><![CDATA[<xsd:choice>]]></c>).
/// <example>
/// <code><![CDATA[
///   <xsd:element name="field" maxOccurs="unbounded" type="CT_Field"/>
/// ]]></code>
/// </example>
/// </summary>
public class ElementType : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    /// <summary>
    /// Name of the element in XML.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// A reference to a <see cref="ComplexType"/>.
    /// </summary>
    public required string TypeName { get; init; }

    public required Occurrences Occurrences { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal void Generate(CodeBuilder code, string namespaceField)
    {
        var typeName = code.NormalizeCt(TypeName);
        var elementParseCall = $"Parse{typeName}(\"{Name}\");";
        var openArgs = $"\"{Name}\", {namespaceField}";
        var min = Occurrences.Min ?? 1;
        var max = Occurrences.Max ?? 1;

        if (min == 1 && max == 1)
        {
            code.AddLine($"_reader.Open({openArgs}))")
                .AddLine(elementParseCall);
        }
        else if (min == 0 && max == 1)
        {
            code.AddLine($"if (_reader.TryOpen({openArgs}))")
                .OpenBrace()
                .AddLine(elementParseCall)
                .CloseBrace();
        }
        else if (min == 0 && max == int.MaxValue)
        {
            code.AddLine($"while (_reader.TryOpen({openArgs}))")
                .OpenBrace()
                .AddLine(elementParseCall)
                .CloseBrace();
        }
        else if (min == 1 && max == int.MaxValue)
        {
            code.AddLine($"_reader.Open({openArgs});")
                .AddLine("do")
                .OpenBrace()
                .AddLine(elementParseCall)
                .CloseBrace()
                .AddLine($"while (_reader.TryOpen({openArgs}));");
        }
        else
        {
            throw new NotSupportedException($"Unexpected occurence range {min}-{max}.");
        }
    }
}
