using ClosedXML.IO.CodeGen.Model.Elements;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:choice>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_Tables">
///   <xsd:choice minOccurs="1" maxOccurs="unbounded">
///     <xsd:element name="m" type="CT_TableMissing"/>
///     <xsd:element name="s" type="CT_XStringElement"/>
///   </xsd:choice>
///   <xsd:attribute name="count" use="optional" type="xsd:unsignedInt"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class ComplexTypeChoice : ComplexType
{
    public required List<IElementGroup> Choices { get; init; } = [];
}
