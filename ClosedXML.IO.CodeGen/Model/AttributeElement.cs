namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A representation of a <c>attribute</c> element in <c>xsd</c> file.
/// <code><![CDATA[
/// <xsd:attribute name="level" type="xsd:unsignedInt" use="optional" default="0"/>
/// ]]></code>
/// </summary>
public class AttributeElement
{
    /// <summary>
    /// Name is technically optional in ref attribute:
    /// <code>
    ///   <![CDATA[<xsd:attribute ref="r:id" use="optional"/>]]>
    /// </code>
    /// </summary>
    public required string? Name { get; set; }

    public required string? Ref { get; set; }

    public required string? Type { get; set; }

    public AttributeUseType Use { get; set; } = AttributeUseType.Optional;

    public string? DefaultValue { get; set; }
}
