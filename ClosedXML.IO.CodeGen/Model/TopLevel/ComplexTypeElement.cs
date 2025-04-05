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
}
