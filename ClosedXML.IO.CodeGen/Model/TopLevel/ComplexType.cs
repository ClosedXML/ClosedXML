using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// Base class for nodes representing a <c><![CDATA[<xsd:compleType>]]></c>.
/// </summary>
public abstract class ComplexType : IReferencable
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required string Name { get; set; }

    public List<OneOf<AttributeElement, AttributeGroupReference>> Attributes { get; set; } = [];

    /// <summary>
    /// Can text be freely interspersed with elements? Only used when <c>complexType</c> contains
    /// <c>any</c>.
    /// </summary>
    public required bool? Mixed { get; init; }

    internal void Generate(CodeBuilder code, string namespaceField)
    {
        code.StartMethod("void Parse{0}(string elementName)", Name);
        code.OpenBrace();
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                attribute.Generate(code);
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        GenerateParseMethod(code, namespaceField);
        CallListener(code);
        code.CloseBrace();

        AddPartialMethodSignature(code, Name);
    }

    internal abstract void GenerateParseMethod(CodeBuilder code, string namespaceField);

    private void CallListener(CodeBuilder code)
    {
        code.WriteIndent().Append("On").AppendComplexType(Name).Append("Parsed(");
        var isFirst = true;
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                if (!isFirst)
                    code.Append(", ");
                code.AppendVariable(attribute.Name!);
                isFirst = false;
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        code.Append(");").EndLine();
    }

    private void AddPartialMethodSignature(CodeBuilder code, string typeName)
    {
        code.EndLine();
        code.WriteIndent().Append($"partial void On").AppendComplexType(typeName).Append("Parsed(");

        var isFirst = true;
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                if (!isFirst)
                    code.Append(", ");

                code.AppendSimpleType(attribute).Append(" ").AppendVariable(attribute.Name!);
                isFirst = false;
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        code.Append(");").EndLine();
    }
}
