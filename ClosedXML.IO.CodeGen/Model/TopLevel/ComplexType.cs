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

        // Get return type of Parse* method
        if (!code.TryGetComplexType(Name, out var csReturnType))
            csReturnType = "void";

        code.StartMethod($"{csReturnType} Parse{{0}}(string elementName)", Name);
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

        if (csReturnType == "void")
        {
            code.WriteIndent();
            AppendCallListener(code, dataVariables);
            code.CloseBrace();

            code.EndLine();
            code.AppendHookSignature(Name, dataVariables);
        }
        else
        {
            code.WriteIndent().Append("return ");
            AppendCallListener(code, dataVariables);
            code.CloseBrace();
        }
    }

    internal abstract List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField);

    private void AppendCallListener(CodeBuilder code, IReadOnlyList<Variable> arguments)
    {
        code.Append("On").AppendComplexType(Name).Append("Parsed(");
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
}
