namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Length of a <see cref="ISimpleType"/> must have a specified length.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="xsd:hexBinary">
///     <xsd:length value="4"/>
///   </xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
public record RestrictLength(int Value) : IValueRestriction;
