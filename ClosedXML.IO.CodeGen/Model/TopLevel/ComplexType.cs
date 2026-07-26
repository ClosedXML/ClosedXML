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

    void IParslet.GenerateParseMethod(CodeBuilder code)
    {
        var attributeVariables = new List<Variable>();
        var csReturnType = code.StartParseMethod(Name, "string elementName", "string ns");
        code.OpenBrace();

        // If we are not in correct element, don't continue onwards.
        code.WriteIndent().Append("if (!_reader.TryOpen(elementName, ns))").EndLine();
        code.OpenBrace();
        if (csReturnType is not null)
        {
            code.AddLine($"return Xpr.Fail<{csReturnType}>();");
        }
        else
        {
            code.AddLine("return Xpr.Fail();");
        }
        code.CloseBrace();
        code.EndLine();

        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                var attributeVariable = attribute.Generate(code);
                attributeVariables.Add(attributeVariable);
            }
            else
            {
                // There are only 2 attribute groups, just hand-code it manually in the reader.
                var agParseMethod = "Parse" + attributeGroup.RefName[3..];
                if (code.TryGetCsType(attributeGroup.RefName, out var csType))
                {
                    var agVarName = char.ToLowerInvariant(attributeGroup.RefName[3]) + attributeGroup.RefName[4..];
                    code.AddLine($"{csType} {agVarName} = {agParseMethod}();");
                    attributeVariables.Add(new Variable(csType, agVarName));
                }
                else
                {
                    code.AddLine($"{agParseMethod}();");
                }
            }
        }

        if (Attributes.Count > 0)
            code.EndLine();

        var elementVariables = GenerateParseMethod(code);
        List<Variable> dataVariables = [.. elementVariables, .. attributeVariables];

        code.AddLine("_reader.Close(elementName, ns);");
        code.EndLine();

        if (csReturnType is null)
        {
            code.WriteIndent().AppendCallHook(Name, dataVariables).Append(";").EndLine();
            code.AddLine("return Xpr.Success();");
            code.CloseBrace();
            code.EndLine();
            code.AddHookSignature(Name, dataVariables);
        }
        else
        {
            // If the Parse* method should map to a value, it's not possible to use partial hook.
            // Partial methods can't return value. The method will be displayed as uncompilable,
            // which is desirable, so it is implemented in the partial reader class by the developer.
            code.WriteIndent().Append("return Xpr.From(").AppendCallHook(Name, dataVariables).Append(");").EndLine();
            code.CloseBrace();
        }
    }

    internal abstract List<Variable> GenerateParseMethod(CodeBuilder code);
}
