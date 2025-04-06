using System.Collections.Generic;
using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.SimpleTypes;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// Parser to parse XSD of OOXML. It doesn't have to support anythings not found in the official XSD.
/// </summary>
public class XsdSchemaParser
{
    /// <summary>
    /// XSD namespace.
    /// </summary>
    private const string XsdNs = "http://www.w3.org/2001/XMLSchema";

    public Schema ParseSchema(XmlTreeReader reader)
    {
        var file = new Schema();

        reader.Open("schema", XsdNs);

        // Read imports
        while (reader.TryOpen("import", XsdNs))
        {
            var ns = reader.GetString("namespace");
            var schemaLocation = reader.GetString("schemaLocation");
            reader.Close("import", XsdNs);

            file.Imports.Add(new ImportElement
            {
                Namespace = ns,
                SchemaLocation = schemaLocation
            });
        }

        while (!reader.TryClose("schema", XsdNs))
        {
            if (reader.TryOpen("complexType", XsdNs))
            {
                var complexType = ParseComplexType(reader);
                file.Entries.Add(complexType);
            }
            else if (reader.TryOpen("simpleType", XsdNs))
            {
                var simpleType = ParseSimpleType(reader);
                file.Entries.Add(simpleType);
            }
            else if (reader.TryOpen("element", XsdNs))
            {
                var name = reader.GetString("name");
                var typeName = reader.GetString("type");
                reader.Close("element", XsdNs);

                file.Entries.Add(new ElementDefinition
                {
                    Name = name,
                    TypeName = typeName
                });
            }
            else if (reader.TryOpen("group", XsdNs))
            {
                var name = reader.GetString("name");
                var elementGroup = ParseElementsGroup(reader);
                reader.Close("group", XsdNs);

                file.Entries.Add(new GroupDefinition
                {
                    Name = name,
                    Content = elementGroup
                });
            }
            else if (reader.TryOpen("attributeGroup", XsdNs))
            {
                var attributeGroup = ParseAttributeGroupDefinition(reader);
                file.Entries.Add(attributeGroup);
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return file;
    }

    private static ComplexType ParseComplexType(XmlTreeReader reader)
    {
        var name = reader.GetString("name");
        var mixed = reader.GetOptionalBool("mixed");
        if (reader.TryOpen("sequence", XsdNs))
        {
            var occurrences = GetOccursAttributes(reader);
            var elements = new List<IElementGroup>();
            do
            {
                var element = ParseElementsGroup(reader);
                elements.Add(element);
            } while (!reader.TryClose("sequence", XsdNs));

            var attributes = ParseComplexTypeAttributes(reader);

            return new ComplexTypeSequence
            {
                Name = name,
                Attributes = attributes,
                Mixed = mixed,
                Sequence = new Sequence
                {
                    Children = elements,
                    Occurrences = occurrences
                }
            };
        }

        if (reader.TryOpen("choice", XsdNs))
        {
            var occurrences = GetOccursAttributes(reader);
            var choices = new List<IElementGroup>();
            do
            {
                var elementGroup = ParseElementsGroup(reader);
                choices.Add(elementGroup);
            } while (!reader.TryClose("choice", XsdNs));

            var attributes = ParseComplexTypeAttributes(reader);

            return new ComplexTypeChoice
            {
                Name = name,
                Attributes = attributes,
                Mixed = mixed,
                Choice = new Choice
                {
                    Children = choices,
                    Occurrences = occurrences
                }
            };
        }

        if (reader.TryOpen("simpleContent", XsdNs))
        {
            var (baseTypeName, extensionAttributes) = ParseSimpleContent(reader);
            var attributes = ParseComplexTypeAttributes(reader);

            return new ComplexTypeSimpleContent
            {
                Name = name,
                Attributes = attributes,
                Mixed = mixed,
                BaseTypeName = baseTypeName,
                ExtensionAttributes = extensionAttributes
            };
        }

        // Complex type that consists only from attributes
        var attr = ParseComplexTypeAttributes(reader);
        return new ComplexTypeElement
        {
            Name = name,
            Attributes = attr,
            Mixed = mixed,
        };
    }

    private static ISimpleType ParseSimpleType(XmlTreeReader reader)
    {
        var simpleTypeName = reader.GetString("name");
        if (reader.TryOpen("restriction", XsdNs))
        {
            var restriction = ParseRestriction(reader);
            reader.Close("simpleType", XsdNs);

            return new SimpleType
            {
                Name = simpleTypeName,
                BaseTypeName = restriction.BaseTypeName,
                Restrictions = restriction.ValueRestrictions,
            };
        }

        if (reader.TryOpen("list", XsdNs))
        {
            var itemType = reader.GetString("itemType");
            reader.Close("list", XsdNs);
            reader.Close("simpleType", XsdNs);

            return new SimpleTypeList
            {
                Name = simpleTypeName,
                ItemType = itemType
            };
        }

        if (reader.TryOpen("union", XsdNs))
        {
            var unionRestrictions = new List<Restriction>();
            while (reader.TryOpen("simpleType", XsdNs))
            {
                reader.Open("restriction", XsdNs);
                var restriction = ParseRestriction(reader);
                reader.Close("simpleType", XsdNs);

                unionRestrictions.Add(restriction);
            }

            reader.Close("union", XsdNs);
            reader.Close("simpleType", XsdNs);

            return new SimpleTypeUnion
            {
                Name = simpleTypeName,
                RestrictionsUnion = unionRestrictions
            };
        }

        throw PartStructureException.ExpectedChoiceElementNotFound(reader);
    }

    private static Restriction ParseRestriction(XmlTreeReader reader)
    {
        var baseType = reader.GetString("base");
        var valueRestrictions = new List<IValueRestriction>();

        while (!reader.TryClose("restriction", XsdNs))
        {
            if (reader.TryOpen("enumeration", XsdNs))
            {
                var value = reader.GetString("value");
                valueRestrictions.Add(new RestrictEnumeration(value));
                reader.Close("enumeration", XsdNs);
            }
            else if (reader.TryOpen("length", XsdNs))
            {
                var length = reader.GetInt("value");
                valueRestrictions.Add(new RestrictLength(length));
                reader.Close("length", XsdNs);
            }
            else if (reader.TryOpen("minInclusive", XsdNs))
            {
                var minInclusive = reader.GetInt("value");
                valueRestrictions.Add(new RestrictMinInclusive(minInclusive));
                reader.Close("minInclusive", XsdNs);
            }
            else if (reader.TryOpen("maxInclusive", XsdNs))
            {
                var maxInclusive = reader.GetInt("value");
                valueRestrictions.Add(new RestrictMaxInclusive(maxInclusive));
                reader.Close("maxInclusive", XsdNs);
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return new Restriction
        {
            BaseTypeName = baseType,
            ValueRestrictions = valueRestrictions
        };
    }

    private static AttributeGroupDefinition ParseAttributeGroupDefinition(XmlTreeReader reader)
    {
        var name = reader.GetString("name");
        var attributes = new List<AttributeElement>();

        while (reader.TryOpen("attribute", XsdNs))
        {
            var attribute = ParseAttribute(reader);
            attributes.Add(attribute);
        }

        reader.Close("attributeGroup", XsdNs);

        return new AttributeGroupDefinition
        {
            Name = name,
            Attributes = attributes
        };
    }

    private static (string Base, List<AttributeElement> Attributes) ParseSimpleContent(XmlTreeReader reader)
    {
        reader.Open("extension", XsdNs);
        var baseTypeName = reader.GetString("base");
        var extensionAttributes = new List<AttributeElement>();

        while (!reader.TryClose("extension", XsdNs))
        {
            reader.Open("attribute", XsdNs);
            var name = reader.GetString("name");
            var type = reader.GetString("type");
            var use = reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Default;
            var defaultValue = reader.GetOptionalString("default");
            reader.Close("attribute", XsdNs);
            var attribute = new AttributeElement
            {
                Name = name,
                Type = type,
                Use = use,
                DefaultValue = defaultValue,
                RefName = null
            };
            extensionAttributes.Add(attribute);
        }

        reader.Close("simpleContent", XsdNs);

        return (baseTypeName, extensionAttributes);
    }

    private static List<OneOf<AttributeElement, AttributeGroupReference>> ParseComplexTypeAttributes(XmlTreeReader reader)
    {
        var attributes = new List<OneOf<AttributeElement, AttributeGroupReference>>();

        while (!reader.TryClose("complexType", XsdNs))
        {
            if (reader.TryOpen("attribute", XsdNs))
            {
                var attribute = ParseAttribute(reader);
                attributes.Add(attribute);
            }
            else if (reader.TryOpen("attributeGroup", XsdNs))
            {
                var refName = reader.GetString("ref");
                reader.Close("attributeGroup", XsdNs);
                attributes.Add(new AttributeGroupReference
                {
                    RefName = refName
                });
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return attributes;
    }

    private static AttributeElement ParseAttribute(XmlTreeReader reader)
    {
        var name = reader.GetOptionalString("name");
        var type = reader.GetOptionalString("type");
        var refName = reader.GetOptionalString("ref");
        var defaultValue = reader.GetOptionalString("default");
        var use = reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Default;
        reader.Close("attribute", XsdNs);

        return new AttributeElement
        {
            Name = name,
            RefName = refName,
            Type = type,
            Use = use,
            DefaultValue = defaultValue
        };
    }

    private static IElementGroup ParseElementsGroup(XmlTreeReader reader)
    {
        if (reader.TryOpen("sequence", XsdNs))
        {
            var occurs = GetOccursAttributes(reader);
            var elements = new List<IElementGroup>();
            do
            {
                var element = ParseElementsGroup(reader);
                elements.Add(element);
            } while (!reader.TryClose("sequence", XsdNs));

            return new Sequence
            {
                Children = elements,
                Occurrences = occurs
            };
        }

        if (reader.TryOpen("choice", XsdNs))
        {
            var occurs = GetOccursAttributes(reader);
            var choices = new List<IElementGroup>();
            do
            {
                var choice = ParseElementsGroup(reader);
                choices.Add(choice);
            } while (!reader.TryClose("choice", XsdNs));

            return new Choice
            {
                Children = choices,
                Occurrences = occurs
            };
        }

        if (reader.TryOpen("element", XsdNs))
        {
            var occurrences = GetOccursAttributes(reader);

            var refName = reader.GetOptionalString("ref");
            if (refName is not null)
            {
                reader.Close("element", XsdNs);

                return new ElementReference
                {
                    RefName = refName,
                    Occurrences = occurrences
                };
            }

            // name, type, min/maxOccurs
            var name = reader.GetString("name");
            var type = reader.GetString("type");
            reader.Close("element", XsdNs);

            return new ElementType
            {
                Name = name,
                TypeName = type,
                Occurrences = occurrences
            };
        }

        if (reader.TryOpen("group", XsdNs))
        {
            var refName = reader.GetOptionalString("ref");
            var occurrences = GetOccursAttributes(reader);

            // Element group reference
            if (refName is not null)
            {
                reader.Close("group", XsdNs);
                return new GroupReference
                {
                    RefName = refName,
                    Occurrences = occurrences
                };
            }

            throw PartStructureException.InvalidAttributeValue();
        }

        if (reader.TryOpen("any", XsdNs))
        {
            var processContents = reader.GetOptionalEnum<ProcessContents>("processContents") ?? ProcessContents.Strict;
            reader.Close("any", XsdNs);

            return new Any
            {
                ProcessContent = processContents
            };
        }

        throw PartStructureException.ExpectedChoiceElementNotFound(reader);
    }

    private static Occurrences GetOccursAttributes(XmlTreeReader reader)
    {
        var minOccurs = reader.GetOptionalInt("minOccurs") ?? null;
        var maxOccurs = reader.GetOptionalString("maxOccurs") == "unbounded" ? int.MaxValue : reader.GetOptionalInt("maxOccurs") ?? null;
        return new Occurrences(minOccurs, maxOccurs);
    }
}
