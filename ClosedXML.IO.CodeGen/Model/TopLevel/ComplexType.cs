using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// Base class for nodes representing a <c><![CDATA[<xsd:compleType>]]></c>.
/// </summary>
public abstract class ComplexType : IParslet
{
    /// <summary>
    /// Name of the complex type.
    /// </summary>
    public required ParsletName Name { get; set; }

    public List<OneOf<AttributeElement, AttributeGroupReference>> Attributes { get; set; } = [];

    /// <summary>
    /// Can text be freely interspersed with elements? Only used when <c>complexType</c> contains
    /// <c>any</c>.
    /// </summary>
    public required bool? Mixed { get; init; }

    void IParslet.GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        var attributeVariables = new List<Variable>();
        var csReturnType = code.StartParseMethod(Name, "string elementName");
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
            code.WriteIndent().AppendCallHook(Name, dataVariables).Append(";").EndLine();
            code.CloseBrace();
            code.EndLine();
            code.AddHookSignature(Name, dataVariables);
        }
        else
        {
            code.WriteIndent().Append("return ").AppendCallHook(Name, dataVariables).Append(";").EndLine();
            code.CloseBrace();
        }
    }

    internal abstract List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField);
}
