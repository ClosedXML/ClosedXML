using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:group/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:group name="EG_ExtensionList" >
///   <xsd:sequence>
///     <xsd:element name = "ext" type="CT_Extension" minOccurs="0" maxOccurs="unbounded"/>
///   </xsd:sequence>
/// </xsd:group>
/// ]]></code>
/// </example>
/// </summary>
public class GroupDefinition : IReferencable
{
    public required string Name { get; init; }

    public required IElementGroup Content { get; init; }
}
