using System;
using System.Diagnostics;
using System.Text;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.SimpleTypes;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

internal class XsdCopyVisitor : IXsdVisitor<Unit>
{
    private readonly StringBuilder _sb;
    private int _indent = 0;

    public XsdCopyVisitor(StringBuilder sb)
    {
        _sb = sb;
    }

    public Unit Visit(Schema schema)
    {
        _sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
        AppendElement("""
                      <xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                                  xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                                  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
                                  xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                                  xmlns:s="http://purl.oclc.org/ooxml/officeDocument/sharedTypes" 
                                  targetNamespace="http://purl.oclc.org/ooxml/spreadsheetml/main"
                                  elementFormDefault="qualified">
                      """);
        foreach (var import in schema.Imports)
        {
            AppendElement($"<xsd:import namespace=\"{import.Namespace}\" schemaLocation=\"{import.SchemaLocation}\"/>");
        }

        foreach (var entry in schema.Entries)
        {
            switch (entry)
            {
                case ComplexTypeSequence ct:
                    ct.Accept(this);
                    break;
                case ComplexTypeChoice ct:
                    ct.Accept(this);
                    break;
                case ComplexTypeSimpleContent ct:
                    ct.Accept(this);
                    break;
                case ComplexTypeElement ct:
                    ct.Accept(this);
                    break;
                case SimpleType simpleType:
                    simpleType.Accept(this);
                    break;
                case SimpleTypeList simpleType:
                    simpleType.Accept(this);
                    break;
                case SimpleTypeUnion simpleType:
                    simpleType.Accept(this);
                    break;
                case ElementDefinition elementDefinition:
                    elementDefinition.Accept(this);
                    break;
                case GroupDefinition groupDefinition:
                    groupDefinition.Accept(this);
                    break;
                case AttributeGroupDefinition attributeGroupDefinition:
                    attributeGroupDefinition.Accept(this);
                    break;
                default:
                    throw new UnreachableException();
            }
        }

        AppendElement("</xsd:schema>");
        return Unit.Value;
    }

    public Unit Visit(ComplexTypeSequence complexType)
    {
        var element = $"<xsd:complexType name=\"{complexType.Name}\"";
        if (complexType.Mixed is not null)
            element += $" mixed=\"{(complexType.Mixed.Value ? "true" : "false")}\"";

        element += ">";
        AppendElement(element);
        complexType.Sequence.Accept(this);
        WriteAttributes(complexType);
        AppendElement("</xsd:complexType>");
        return Unit.Value;
    }

    public Unit Visit(ComplexTypeChoice complexType)
    {
        AppendElement($"<xsd:complexType name=\"{complexType.Name}\">");
        complexType.Choice.Accept(this);
        WriteAttributes(complexType);
        AppendElement("</xsd:complexType>");
        return Unit.Value;
    }

    public Unit Visit(ComplexTypeSimpleContent complexType)
    {
        AppendElement($"<xsd:complexType name=\"{complexType.Name}\">");
        AppendElement("<xsd:simpleContent>");
        AppendElement($"<xsd:extension base=\"{complexType.BaseTypeName}\">");
        WriteAttributes(complexType);
        AppendElement("</xsd:extension>");
        AppendElement("</xsd:simpleContent>");
        AppendElement("</xsd:complexType>");
        return Unit.Value;
    }

    public Unit Visit(ComplexTypeElement complexType)
    {
        AppendElement($"<xsd:complexType name=\"{complexType.Name}\">");
        WriteAttributes(complexType);
        AppendElement("</xsd:complexType>");
        return Unit.Value;
    }

    public Unit Visit(AttributeGroupReference attributeGroupReference)
    {
        AppendElement($"<xsd:attributeGroup ref=\"{attributeGroupReference.RefName}\"/>");
        return Unit.Value;
    }

    public Unit Visit(SimpleType simpleType)
    {
        AppendElement($"<xsd:simpleType name=\"{simpleType.Name}\">");
        AppendElement($"<xsd:restriction base=\"{simpleType.BaseTypeName}\">");

        foreach (var restriction in simpleType.Restrictions)
            AppendElement(GetValueRestrictionElement(restriction));

        AppendElement("</xsd:restriction>");
        AppendElement("</xsd:simpleType>");
        return Unit.Value;
    }

    private static string GetValueRestrictionElement(IValueRestriction restriction)
    {
        var restrictionElement = restriction switch
        {
            RestrictEnumeration enumeration => $"<xsd:enumeration value=\"{enumeration.Value}\"/>",
            RestrictLength length => $"<xsd:length value=\"{length.Value}\"/>",
            RestrictMinInclusive minInclusive => $"<xsd:minInclusive value=\"{minInclusive.Value}\"/>",
            RestrictMaxInclusive maxInclusive => $"<xsd:maxInclusive value=\"{maxInclusive.Value}\"/>",
            _ => throw new UnreachableException()
        };
        return restrictionElement;
    }

    public Unit Visit(SimpleTypeList simpleType)
    {
        AppendElement($"<xsd:simpleType name=\"{simpleType.Name}\">");
        AppendElement($"<xsd:list itemType=\"{simpleType.ItemType}\"/>");
        AppendElement("</xsd:simpleType>");
        return Unit.Value;
    }

    public Unit Visit(SimpleTypeUnion simpleType)
    {
        AppendElement($"<xsd:simpleType name=\"{simpleType.Name}\">");
        AppendElement("<xsd:union>");
        foreach (var restrictionUnion in simpleType.RestrictionsUnion)
        {
            AppendElement("<xsd:simpleType>");
            AppendElement($"<xsd:restriction base=\"{restrictionUnion.BaseTypeName}\">");
            foreach (var valueRestriction in restrictionUnion.ValueRestrictions)
            {
                AppendElement(GetValueRestrictionElement(valueRestriction));
            }

            AppendElement("</xsd:restriction>");
            AppendElement("</xsd:simpleType>");
        }

        AppendElement("</xsd:union>");
        AppendElement("</xsd:simpleType>");
        return Unit.Value;
    }

    public Unit Visit(GroupReference attributeGroupReference)
    {
        AppendElement($"<xsd:group ref=\"{attributeGroupReference.RefName}\"{GetOccurrences(attributeGroupReference.Occurrences)}/>");
        return Unit.Value;
    }

    public Unit Visit(Any any)
    {
        var element = "<xsd:any";
        element += any.ProcessContent switch
        {
            ProcessContents.Default => string.Empty,
            ProcessContents.Strict => " processContents=\"strict\"",
            ProcessContents.Lax => " processContents=\"lax\"",
            _ => throw new UnreachableException(),
        };
        element += "/>";
        AppendElement(element);
        return Unit.Value;
    }

    public Unit Visit(Choice choice)
    {
        AppendElement($"<xsd:choice{GetOccurrences(choice.Occurrences)}>");
        foreach (var element in choice.Children)
        {
            element.Accept(this);
        }
        AppendElement("</xsd:choice>");
        return Unit.Value;
    }

    public Unit Visit(ElementType elementType)
    {
        var element = $"<xsd:element name=\"{elementType.Name}\" type=\"{elementType.TypeName}\"";
        element += GetOccurrences(elementType.Occurrences);
        element += "/>";
        AppendElement(element);
        return Unit.Value;
    }

    public Unit Visit(Sequence sequence)
    {
        AppendElement($"<xsd:sequence{GetOccurrences(sequence.Occurrences)}>");
        foreach (var element in sequence.Children)
        {
            element.Accept(this);
        }
        AppendElement("</xsd:sequence>");
        return Unit.Value;
    }

    public Unit Visit(ElementReference elementReference)
    {
        AppendElement($"<xsd:element ref=\"{elementReference.RefName}\"{GetOccurrences(elementReference.Occurrences)}/>");
        return Unit.Value;
    }

    public Unit Visit(ElementDefinition elementDefinition)
    {
        AppendElement($"<xsd:element name=\"{elementDefinition.Name}\" type=\"{elementDefinition.TypeName}\"/>");
        return Unit.Value;
    }

    public Unit Visit(GroupDefinition groupDefinition)
    {
        AppendElement($"<xsd:group name=\"{groupDefinition.Name}\">");
        groupDefinition.Content.Accept(this);
        AppendElement("</xsd:group>");
        return Unit.Value;
    }

    public Unit Visit(AttributeGroupDefinition attributeGroupDefinition)
    {
        AppendElement($"<xsd:attributeGroup name=\"{attributeGroupDefinition.Name}\">");
        foreach (var attribute in attributeGroupDefinition.Attributes)
            attribute.Accept(this);

        AppendElement("</xsd:attributeGroup>");
        return Unit.Value;
    }

    public Unit Visit(AttributeElement attributeElement)
    {
        var element = "<xsd:attribute";
        element += attributeElement.RefName is not null
            ? $" ref=\"{attributeElement.RefName}\""
            : $" name=\"{attributeElement.Name}\" type=\"{attributeElement.Type}\"";

        if (attributeElement.Use != AttributeUseType.Default)
            element += $" use=\"{(attributeElement.Use == AttributeUseType.Optional ? "optional" : "required")}\"";

        if (attributeElement.DefaultValue is not null)
            element += $" default=\"{attributeElement.DefaultValue}\"";

        element += "/>";
        AppendElement(element);
        return Unit.Value;
    }

    private void WriteAttributes(ComplexType complexType)
    {
        foreach (var attr in complexType.Attributes)
        {
            if (attr.TryPickT1(out var attribute))
            {
                attribute.Accept(this);
            }
            else if (attr.TryPickT2(out var attributeGroup))
            {
                attributeGroup.Accept(this);
            }
            else
            {
                throw new InvalidOperationException();
            }
        }
    }

    private static string GetOccurrences(Occurrences occurrences)
    {
        var attributes = string.Empty;
        if (occurrences.Min is not null)
            attributes += $" minOccurs=\"{occurrences.Min}\"";

        if (occurrences.Max is not null)
            attributes += $" maxOccurs=\"{(occurrences.Max == int.MaxValue ? "unbounded" : occurrences.Max.ToString())}\"";

        return attributes;
    }

    private void AppendElement(string text)
    {
        _sb.Append(' ', _indent);
        _sb.AppendLine(text);

        const StringComparison comparison = StringComparison.Ordinal;
        var isOpen = text.StartsWith('<') && !text.StartsWith("</", comparison);
        if (isOpen)
            _indent += 2;

        var isClose = text.StartsWith("</", comparison) || text.EndsWith("/>", comparison);
        if (isClose)
            _indent -= 2;
    }
}
