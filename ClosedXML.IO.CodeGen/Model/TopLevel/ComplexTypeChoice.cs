using System;
using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model.Elements;

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
    public required Choice Choice { get; init; }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code)
    {
        var choicesCount = Choice.DetermineChoicesCount();
        switch (choicesCount)
        {
            case ElementsCount.ZeroToOne:
            {
                var variables = Choice.GenerateParseContent(Name, choicesCount, code, true);
                return variables;
            }
            case ElementsCount.ZeroToMany:
            {
                var variables = Choice.GenerateParseContent(Name, choicesCount, code, true);
                return variables;
            }
            case ElementsCount.OneToOne:
            {
                var variables = Choice.GenerateParseContent(Name, choicesCount, code, true);
                return variables;
            }
            case ElementsCount.OneToMany:
            {
                var variables = Choice.GenerateParseContent(Name, choicesCount, code, true);
                return variables;
            }
            default:
                throw new NotImplementedException();
        }
    }
}
