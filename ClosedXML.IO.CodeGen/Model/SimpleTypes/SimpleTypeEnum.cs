using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// <c><![CDATA[<xsd:simpleType>]]></c> inside <c><![CDATA[<xsd:schema>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:simpleType name="ST_Something">
///   <xsd:restriction base="xsd:string">
///     <xsd:enumeration value="equal"/>
///     <xsd:enumeration value="lessThan"/>
///   </xsd:restriction>
/// </xsd:simpleType>
/// ]]></code>
/// </example>
/// </summary>
public class SimpleTypeEnum : ISimpleType
{
    public required string Name { get; init; }

    public required string BaseTypeName { get; init; }

    /// <summary>
    /// List of enum values.
    /// </summary>
    public required List<string> Values { get; init; }

    public required int? Length { get; init; }

    public required int? MinInclusive { get; init; }

    public required int? MaxInclusive { get; init; }
}
