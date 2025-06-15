namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// Value must match regular expression.
/// <example>
/// <code><![CDATA[
///   <xsd:restriction base="xsd:string">
///     <xsd:pattern value="0*((2[5-9])|([3-9][0-9])|([1-3][0-9][0-9])|400)%"/>
///   </xsd:restriction>
/// ]]></code>
/// </example>
/// </summary>
public record RestrictPattern(string Value) : IValueRestriction;
