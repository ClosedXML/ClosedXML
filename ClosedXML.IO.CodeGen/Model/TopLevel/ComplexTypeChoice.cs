using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:choice>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_Tables">
///   <xsd:choice minOccurs="1" maxOccurs="unbounded">
///     <xsd:element name="m" type="CT_TableMissing"/>
///     <xsd:element name="s" type="CT_XStringElement"/>
///   </xsd:choice>
///   <xsd:attribute name="count" use="optional" type="xsd:unsignedInt"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class ComplexTypeChoice : ComplexType, INode
{
    public required Choice Choice { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        var choicesCount = Choice.DetermineChoicesCount();
        switch (choicesCount)
        {
            case ElementsCount.ZeroToOne:
            {
                var variables = Choice.GenerateParseContent(choicesCount, code, namespaceField);
                code.AddLine($"_reader.Close(elementName, {namespaceField});");
                return variables;
            }
            case ElementsCount.OneToMany:
            {
                code.AddLine("do");
                code.OpenBrace();
                var variables = Choice.GenerateParseContent(choicesCount, code, namespaceField);
                code.CloseBrace();
                code.AddLine($"while (!_reader.TryClose(elementName, {namespaceField}));");
                return variables;
            }
            default:
                throw new NotImplementedException();
        }
    }
}
