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

    internal Variable? Generate(CodeBuilder code, string namespaceField)
    {
        Variable? variable = null;
        var typeName = code.NormalizeCt(TypeName);
        var elementParseCall = $"Parse{typeName}(\"{Name}\")";
        var openArgs = $"\"{Name}\", {namespaceField}";
        var min = Occurrences.Min ?? 1;
        var max = Occurrences.Max ?? 1;

        if (min == 1 && max == 1)
        {
            code.AddLine($"_reader.Open({openArgs});");
            code.WriteIndent();
            if (code.TryGetCsType(TypeName, out var csType))
            {
                variable = new Variable(csType, Name);
                code.Append("var ").AppendVariable(Name).Append(" = ");
            }
            code.Append(elementParseCall).Append(";").EndLine();
        }
        else if (min == 0 && max == 1)
        {
            if (code.TryGetCsType(TypeName, out var csType))
            {
                csType += "?";
                variable = new Variable(csType, Name);
                code.WriteIndent().Append(csType).Append(" ").AppendVariable(Name).Append(" = default;").EndLine();
                code.AddLine($"if (_reader.TryOpen({openArgs}))");
                code.OpenBrace();
                code.WriteIndent().AppendVariable(Name).Append(" = ").Append(elementParseCall).Append(";").EndLine();
                code.CloseBrace();
            }
            else
            {
                code.AddLine($"if (_reader.TryOpen({openArgs}))");
                code.OpenBrace();
                code.WriteIndent().Append(elementParseCall).Append(";").EndLine();
                code.CloseBrace();
            }
        }
        else if (min == 0 && max == int.MaxValue)
        {
            if (code.TryGetCsType(TypeName, out var csType))
            {
                csType = $"List<{csType}>";
                variable = new Variable(csType, Name);
                code.WriteIndent().Append("var ").AppendVariable(variable.Name).Append($" = new {csType}();").EndLine();
                code.AddLine($"while (_reader.TryOpen({openArgs}))");
                code.OpenBrace();
                code.WriteIndent().AppendVariable(variable.Name).Append($".Add({elementParseCall})").Append(";").EndLine();
                code.CloseBrace();
            }
            else
            {
                code.AddLine($"while (_reader.TryOpen({openArgs}))");
                code.OpenBrace();
                code.WriteIndent().Append(elementParseCall).Append(";").EndLine();
                code.CloseBrace();
            }
        }
        else if (min == 1 && max == int.MaxValue)
        {
            if (code.TryGetCsType(TypeName, out var csType))
            {
                csType = $"List<{csType}>";
                variable = new Variable(csType, Name);
                code.WriteIndent().Append("var ").AppendVariable(variable.Name).Append($" = new {csType}();").EndLine();
                code.AddLine($"_reader.Open({openArgs});");
                code.AddLine("do");
                code.OpenBrace();
                code.WriteIndent().AppendVariable(variable.Name).Append($".Add({elementParseCall})").Append(";").EndLine();
                code.CloseBrace();
                code.AddLine($"while (_reader.TryOpen({openArgs}));");
            }
            else
            {
                code.AddLine($"_reader.Open({openArgs});");
                code.AddLine("do");
                code.OpenBrace();
                code.WriteIndent().Append(elementParseCall).Append(";").EndLine();
                code.CloseBrace();
                code.AddLine($"while (_reader.TryOpen({openArgs}));");
            }
        }
        else
        {
            throw new NotSupportedException($"Unexpected occurence range {min}-{max}.");
        }

        return variable;
    }
}
