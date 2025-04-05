using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:simpleContent>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <![CDATA[
/// <xsd:complexType name="CT_CellFormula">
///   <xsd:simpleContent>
///     <xsd:extension base="ST_Formula">
///       <xsd:attribute name = "t" type="ST_CellFormulaType" use="optional" default="normal"/>
///       <xsd:attribute name = "aca" type="xsd:boolean" use="optional" default="false"/>
///     </xsd:extension>
///   <xsd:simpleContent>
/// ]]>
/// </summary>
public class ComplexTypeSimpleContent : ComplexType, INode
{
    public required string BaseTypeName { get; init; }

    public required List<AttributeElement> ExtensionAttributes { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
