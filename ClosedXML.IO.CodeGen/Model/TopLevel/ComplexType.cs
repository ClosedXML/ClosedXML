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
        var attributeVariables = new List<Variable>();
        code.StartMethod("void Parse{0}(string elementName)", Name);
        code.OpenBrace();
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                var attributeVariable = attribute.Generate(code);
                attributeVariables.Add(attributeVariable);
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        var elementVariables = GenerateParseMethod(code, namespaceField);
        List<Variable> dataVariables = [.. elementVariables, .. attributeVariables];
        CallListener(code, dataVariables);
        code.CloseBrace();

        AddPartialMethodSignature(code, Name, dataVariables);
    }

    internal abstract List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField);

    private void CallListener(CodeBuilder code, IReadOnlyList<Variable> arguments)
    {
        code.WriteIndent().Append("On").AppendComplexType(Name).Append("Parsed(");
        var isFirst = true;
        foreach (var variable in arguments)
        {
            if (!isFirst)
                code.Append(", ");

            code.AppendVariable(variable.Name);
            isFirst = false;
        }

        code.Append(");").EndLine();
    }

    private void AddPartialMethodSignature(CodeBuilder code, string typeName, IReadOnlyList<Variable> parameters)
    {
        code.EndLine();
        code.WriteIndent().Append("partial void On").AppendComplexType(typeName).Append("Parsed(");

        var isFirst = true;
        foreach (var parameter in parameters)
        {
            if (!isFirst)
                code.Append(", ");

            code.Append(parameter.Type).Append(" ").AppendVariable(parameter.Name);
            isFirst = false;
        }

        code.Append(");").EndLine();
    }
}
