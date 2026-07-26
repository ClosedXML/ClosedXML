using ClosedXML.IO.CodeGen.Model.Elements;
using System;
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
    public required Sequence Sequence { get; init; }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code)
    {
        var dataVariables = new List<Variable>();
        if (Sequence.Occurrences.Elements != ElementsCount.OneToOne)
            throw new NotSupportedException("Only simple sequence is supported. Change XSD structure.");

        // The only sane sequence
        foreach (var element in Sequence.Children)
        {
            if (element is ElementType elementType)
            {
                var variables = elementType.GenerateSequenceParseCode(code);
                dataVariables.AddRange(variables);
            }
            else if (element is GroupReference groupReference)
            {
                var variables = groupReference.GenerateSequenceParseCall(code);
                dataVariables.AddRange(variables);
            }
            else
            {
                throw new NotImplementedException("Only element type is implemented for a sequence.");
            }
        }

        return dataVariables;
    }
}
