using ClosedXML.IO.CodeGen.Model.Elements;
using System;

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

    internal override void GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        var min = Choice.Occurrences.Min ?? 1;
        var max = Choice.Occurrences.Max ?? 1;

        if (min == 1 && max == int.MaxValue)
        {
            code.AddLine("do");
            code.OpenBrace();
            var isFirst = true;
            foreach (var child in Choice.Children)
            {
                var element = (ElementType)child;
                var joiner = isFirst ? string.Empty : "else ";
                isFirst = false;

                code.AddLine($"{joiner}if (_reader.TryOpen(\"{element.Name}\", {namespaceField}))");
                code.OpenBrace();
                code.AddLine($"Parse{code.NormalizeCt(element.TypeName)}(\"{element.Name}\");");
                code.CloseBrace();
            }

            code.AddLine("else");
            code.OpenBrace();
            code.AddLine("throw PartStructureException.ExpectedChoiceElementNotFound(_reader);");
            code.CloseBrace();
            code.CloseBrace();
            code.AddLine($"while (!_reader.TryClose(elementName, {namespaceField}));");
        }
        else
        {
            throw new NotImplementedException($"{min}-{max} choice is not implemented.");
        }
    }
}
