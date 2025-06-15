namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Numerical value must be at least the specified value.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="xsd:integer">
///     <xsd:minInclusive value="0"/>
///     <xsd:maxInclusive value="14"/>
///   </xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
public record RestrictMinInclusive(string Value) : IValueRestriction;
