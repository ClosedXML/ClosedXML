using System;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:group/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:group name="EG_ExtensionList" >
///   <xsd:sequence>
///     <xsd:element name = "ext" type="CT_Extension" minOccurs="0" maxOccurs="unbounded"/>
///   </xsd:sequence>
/// </xsd:group>
/// ]]></code>
/// </example>
/// </summary>
public class GroupDefinition : IParslet, INode
{
    public required string Name { get; init; }

    public required IElementGroup Content { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    void IParslet.GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        if (Content is Choice choice)
        {
            var choicesCount = choice.DetermineChoicesCount();
            if (choicesCount != ElementsCount.OneToOne)
                throw new NotSupportedException("Element group choice should have 1 occurence.");

            var returnCsType = code.StartElementGroupParseMethod(Name);
            code.OpenBrace();
            var variables = choice.GenerateParseContent(choicesCount, code, namespaceField);
            if (returnCsType == "void")
            {
                code.WriteIndent().AppendCallHook(Name, variables);
                code.CloseBrace();
                code.EndLine();
                code.AppendHookSignature(Name, variables);
            }
            else
            {
                code.WriteIndent().Append("return ").AppendCallHook(Name, variables);
                code.CloseBrace();
            }
        }
        else
        {
            throw new NotImplementedException("Only choice implemented.");
        }
    }
}
