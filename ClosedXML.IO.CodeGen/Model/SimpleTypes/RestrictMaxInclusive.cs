namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Numerical value must be at most the specified value.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="xsd:nonNegativeInteger">
///     <xsd:maxInclusive value="180"/>
///   </ xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
public record RestrictMaxInclusive(int Value) : IValueRestriction;
