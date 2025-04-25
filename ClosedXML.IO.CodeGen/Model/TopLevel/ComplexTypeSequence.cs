using ClosedXML.IO.CodeGen.Model.Elements;
using System;

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

    internal override void GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        code.StartMethod("void Parse{0}(string elementName)", Name);
        code.OpenBrace();
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                attribute.Generate(code);
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        var min = Sequence.Occurrences.Min ?? 1;
        var max = Sequence.Occurrences.Max ?? 1;
        if (min == 1 && max == 1)
        {
            foreach (var element in Sequence.Children)
            {
                if (element is ElementType elementType)
                {
                    elementType.Generate(code, namespaceField);
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
        CallListener(code);
        code.CloseBrace();

        AddPartialMethodSignature(code, Name);
    }

    private void CallListener(CodeBuilder code)
    {
        code.WriteIndent().Append("On").AppendComplexType(Name).Append("Parsed(");
        var isFirst = true;
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                if (!isFirst)
                    code.Append(", ");
                code.AppendVariable(attribute.Name!);
                isFirst = false;
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        code.Append(");").EndLine();
    }

    private void AddPartialMethodSignature(CodeBuilder code, string typeName)
    {
        code.EndLine();
        code.WriteIndent().Append($"partial void On").AppendComplexType(typeName).Append("Parsed(");

        var isFirst = true;
        foreach (var oneOfAttribute in Attributes)
        {
            if (oneOfAttribute.TryPickT1(out var attribute, out var attributeGroup))
            {
                if (!isFirst)
                    code.Append(", ");

                code.AppendSimpleType(attribute).Append(" ").AppendVariable(attribute.Name!);
                isFirst = false;
            }
            else
            {
                throw new NotImplementedException($"Attribute group ({attributeGroup.RefName}) not yet implemented.");
            }
        }

        code.Append(");").EndLine();
    }
}
