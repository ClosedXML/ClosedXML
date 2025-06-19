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
public class ComplexTypeSequence : ComplexType, INode
{
    public required Sequence Sequence { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        var dataVariables = new List<Variable>();
        var min = Sequence.Occurrences.Min ?? 1;
        var max = Sequence.Occurrences.Max ?? 1;
        if (min == 1 && max == 1)
        {
            foreach (var element in Sequence.Children)
            {
                if (element is ElementType elementType)
                {
                    var variable = elementType.Generate(code, namespaceField);
                    if (variable is not null)
                        dataVariables.Add(variable);
                }
                else if (element is GroupReference groupReference)
                {
                    var variable = groupReference.GenerateParseCall(code, namespaceField);
                    if (variable is not null)
                        dataVariables.Add(variable);
                }
                else
                {
                    throw new NotImplementedException("Only element type is implemented for a sequence.");
                }
            }
        }
        else
        {
            throw new NotImplementedException("Only simple sequence is implemented.");
        }

        code.AddLine($"_reader.Close(elementName, {namespaceField});");
        return dataVariables;
    }
}
