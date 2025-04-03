using ClosedXML.IO.CodeGen.Model.Elements;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:sequence>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_AutoFilter">
///   <xsd:sequence>
///     <xsd:element name="filterColumn" minOccurs="0" maxOccurs="unbounded" type="CT_FilterColumn"/>
///     <xsd:element name="sortState" minOccurs="0" maxOccurs="1" type="CT_SortState"/>
///   </xsd:sequence>
///   <xsd:attribute name="ref" type="ST_Ref"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class ComplexTypeSequence : ComplexType
{
    public required List<IElementGroup> Elements { get; init; } = [];
}
