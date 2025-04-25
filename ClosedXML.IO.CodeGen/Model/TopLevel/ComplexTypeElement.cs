using System;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>. It doesn't have
/// any elements, only attributes.
/// </summary>
public class ComplexTypeElement : ComplexType, INode
{
    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal override void GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        code.StartMethod("void Parse{0}(string elementName)", Name)
            .OpenBrace();

        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                attribute.Generate(code);
            }
            else
            {
                throw new NotImplementedException($"Attribute group '{attributeGroup}' read not implemented.");
            }
        }
        code.AddLine($"_reader.Close(elementName, {namespaceField});");
        code.CloseBrace();
    }
}
