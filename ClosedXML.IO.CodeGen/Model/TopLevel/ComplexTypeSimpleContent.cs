using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A complex type that is based on a simple type, e.g.,
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
public class ComplexTypeSimpleContent : ComplexType
{
    public required string BaseTypeName { get; init; }

    public required List<AttributeElement> ExtensionAttributes { get; init; }
}
