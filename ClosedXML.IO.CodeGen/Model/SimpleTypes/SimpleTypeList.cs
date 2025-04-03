namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// <c><![CDATA[<xsd:simpleType>]]></c> inside <c><![CDATA[<xsd:schema>]]></c>. List items are
/// separated by a space.
/// <example>
/// <code><![CDATA[
///  <xsd:simpleType name="ST_Sqref">
///    <xsd:list itemType="ST_Ref"/>
///  </xsd:simpleType>
/// ]]></code>
/// </example>
/// </summary>
public class SimpleTypeList
{
    public required string Name { get; init; }

    public required string ItemType { get; init; }
}
