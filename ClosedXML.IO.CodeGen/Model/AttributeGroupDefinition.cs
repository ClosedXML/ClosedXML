namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// Attribute group definition. It is a child of <c>schema</c>.
/// <code><![CDATA[
///  <xsd:attributeGroup name="AG_RevData">
///    <xsd:attribute name="rId" type="xsd:unsignedInt" use="required"/>
///    <xsd:attribute name = "ua" type="xsd:boolean" use="optional" default="false"/>
///    <xsd:attribute name = "ra" type="xsd:boolean" use="optional" default="false"/>
///  </xsd:attributeGroup>
/// ]]></code>
/// </summary>
public class AttributeGroupDefinition : ComplexTypeBase;
