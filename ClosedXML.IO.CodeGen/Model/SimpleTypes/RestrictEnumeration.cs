namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Value of <see cref="ISimpleType"/> can be the <paramref name="Value"/>.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="xsd:string">
///     <xsd:enumeration value="none"/>
///     <xsd:enumeration value="thin"/>
///   </xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
/// <param name="Value">Allowed value of simple type.</param>
public record RestrictEnumeration(string Value) : IValueRestriction;
