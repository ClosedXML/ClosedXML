namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Numerical value must be less than the specified value.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="ST_Angle">
///     <xsd:minExclusive value="-5400000"/>
///     <xsd:maxExclusive value="5400000"/>
///   </xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
public record RestrictMaxExclusive(string Value) : IValueRestriction;
