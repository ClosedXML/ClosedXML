using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// <c><![CDATA[<xsd:simpleType>]]></c> inside <c><![CDATA[<xsd:schema>]]></c>. The allowed values
/// are restricted by <see cref="Restrictions"/>.
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
public class SimpleType : ISimpleType
{
    public required string Name { get; init; }

    public required string BaseTypeName { get; init; }

    /// <summary>
    /// Conditions the value must satisfy.
    /// </summary>
    public required List<IValueRestriction> Restrictions { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}
