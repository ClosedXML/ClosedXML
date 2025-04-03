using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
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
        if (reader.TryOpen("sequence", XsdNs))
        {
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
                Elements = elements
            };
        }

        if (reader.TryOpen("choice", XsdNs))
        {
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
                Choices = choices
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
                BaseTypeName = baseTypeName,
                ExtensionAttributes = extensionAttributes
            };
        }

        // Complex type that consists only from attributes
        var attr = ParseComplexTypeAttributes(reader);
        return new ComplexType
        {
            Name = name,
            Attributes = attr
        };
    }

    private static ISimpleType ParseSimpleType(XmlTreeReader reader)
    {
        var simpleTypeName = reader.GetString("name");
        if (reader.TryOpen("restriction", XsdNs))
        {
            var baseType = reader.GetString("base");

            int? length = null;
            int? minInclusive = null;
            int? maxInclusive = null;
            var values = new List<string>();

            while (!reader.TryClose("restriction", XsdNs))
            {
                if (reader.TryOpen("enumeration", XsdNs))
                {
                    var value = reader.GetString("value");
                    values.Add(value);
                    reader.Close("enumeration", XsdNs);
                }
                else if (reader.TryOpen("length", XsdNs))
                {
                    length = reader.GetInt("value");
                    reader.Close("length", XsdNs);
                }
                else if (reader.TryOpen("minInclusive", XsdNs))
                {
                    minInclusive = reader.GetInt("value");
                    reader.Close("minInclusive", XsdNs);
                }
                else if (reader.TryOpen("maxInclusive", XsdNs))
                {
                    maxInclusive = reader.GetInt("value");
                    reader.Close("maxInclusive", XsdNs);
                }
                else
                {
                    throw PartStructureException.ExpectedChoiceElementNotFound(reader);
                }
            }

            reader.Close("simpleType", XsdNs);

            return new SimpleTypeEnum
            {
                Name = simpleTypeName,
                BaseTypeName = baseType,
                Values = values,
                Length = length,
                MinInclusive = minInclusive,
                MaxInclusive = maxInclusive
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
            // TODO: Implement, but it's a minor use
            reader.Skip();
            reader.Close("simpleType", XsdNs);

            return new SimpleTypeUnion
            {
                Name = simpleTypeName
            };
        }

        throw PartStructureException.ExpectedChoiceElementNotFound(reader);
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
            var use = reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Optional;
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

    private static List<AttributeElement> ParseComplexTypeAttributes(XmlTreeReader reader)
    {
        var attributes = new List<AttributeElement>();

        while (!reader.TryClose("complexType", XsdNs))
        {
            if (reader.TryOpen("attribute", XsdNs))
            {
                var attribute = ParseAttribute(reader);
                attributes.Add(attribute);
            }
            else if (reader.TryOpen("attributeGroup", XsdNs))
            {
                _ = reader.GetString("ref");
                reader.Close("attributeGroup", XsdNs);
                // TODO return XsdAttributeGroupReference, currently ignored
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
        var use = reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Optional;
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
        var minOccurs = reader.GetOptionalInt("minOccurs") ?? 1;
        var maxOccurs = reader.GetOptionalString("maxOccurs") == "unbounded" ? int.MaxValue : reader.GetOptionalInt("maxOccurs") ?? 1;
        return new Occurrences(minOccurs, maxOccurs);
    }
}
