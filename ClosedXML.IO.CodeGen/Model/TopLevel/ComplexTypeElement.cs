using System.Collections.Generic;

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

    internal override List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        code.AddLine($"_reader.Close(elementName, {namespaceField});");
        return [];
    }
}
