# Overview

The goal is to create a generator that will use XSD of OOXML and it will generate parsing logic that includes data extraction and to load extracted data into ClosedXML internal structures.

The data loading part might need to do custom logic that has to be incorporated into the generated parser. There might also be some validation, not just data combination logic.

## Requirements

Generator must
* Be able to generate parsing logic for XSD that extracts data
* Must be able to combine extracted data from generated parser and custom logic/validation
* Must be able to be regenerate parsing code without loss of hand-coded validation/translation logic
* Must be configurable, some parts might be completely generated, some might use hand-coded parser
* Use forward only XML parser `XmlTreeParser`
* Avoid a separate intermediate structure creation
* Must support only XSD features found in OOXML schema, nothing extra needed

## Rationale

Current OpenXML SDK is an intermediate representation that loads each part into memory. That has several problems, the major one is performance, both cpu and memory consumption. OpenXML SDK loads whole part into memory and ClosedXML then reads it and sets internal structures and then the whole parsed XML tree is disposed of. That is slow and memory intensive.

To solve it, we will use our custom parser that is
* forward only
* will handle ISO-29500-3 (Markup Compatibility and Extensibility)
* is designed to be hand-coded

We want to avoid intermediate representation, because that is what we already have. I could try to make one that is more optimal, but I don't see benefit. It would just be extra layer and extra work.

It's inevitable that there will be bugs in the generated code. Bugs must be fixed and fixed everywhere. Therefore regeneration of code without affecting the hand logic is crucial.

# Validation during parsing

The generated code parses the expected schema and validates that the XML conforms to the schema. If the XML doesn't match the schema, the generated parser will throw an exception.

This property ensures that when a hook is called, we can be certain that the XML processed up to that point was valid. This is the key difference between a CodeGen parser and a classic hand-coded `XmlReader` parser, as shown in this example:

```csharp
// Classic XmlReader hand-coded parser. No explicit validation, requires to
// supply schema to XmlReader and set XmlReaderSettings.ValidationType to
// ValidationType.Schema.
while (reader.Read()) {
   if (reader.IsStartElement()) {
     if (reader.Name == "numFmts")
       // Do something
     } else if (reader.Name == "fonts")
       // Do something else
     } else if (reader.Name == "fills") {
       // ...
     }
   }
}
````

Of course CodeGen parser is limited in other ways, but for purposes of OOXML it is a better choice.

# Usage

In order to generate a parser, it is necessary to

* register simple types to the `SchemeTypeMap`
* define type mapping for complex types to C# types (optional)
* define which complex types to generate generated parser

## Simple Type Mapping

Simple type is an XML value type used in the attributes. It defines mapping between XML type and C# type. In most cases, we use primitie types in C#, but any type can be used. Use `AddSimpleType` to register a mapping.

The mapping must contain at least one code fragment template that will map the attribute value to the C# value. There are two possible template:
* RequiredTemplate - a code fragment that maps XML attribute value to C# value (attribute name is supplied thorugh `{0}`). It should should throw an exception when attribute can't be mapped to the value or is missing.
* OptionalTemplate - a code fragment that maps XML attribute value to C# value (attribute name is supplied thorugh `{0}`). The fragment is used when attribute is optional and should return `null` when the attribute is missing.

```csharp
var typeMap = new SchemeTypeMap()
    // Adds some very common simple type mapping used in basically every reader
    .AddPrimitiveTypes()
    .AddSimpleType(new SimpleTypeMapping
    {
        Name = "ST_NumFmtId",
        CsTypeName = "uint",
        RequiredTemplate = "_reader.GetUInt(\"{0}\")",
        OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")"
    })
```

There are few other methods for mapping enum, all start with `AddSimpleType*`.

If CodeGen can't find mapping for a type, it will throw an exception during generation.

## Complex Type Mapping

CodeGen expects that there is a parsing method for each type (e.g. `ParseFont` for type `CT_Font`). By default, the generator expects that each parsing method returns `void`. It can be useful instead to return a value and use the returned value in parent `Parse*` method.

In order to do that, use `AddComplexTypeMapping` in `SchemeTypeMap` to define this mapping.

The mapping doesn't mean that the method will be generated, but that other `Parse*` method will expect that `Parse*` method for the type returns a value.

```csharp
typeMap
  // Code generator will expect that method ParseColor will return XLColor.
  // The actual ParseColor method is not generated, it is hand-coded, but
  // the generated ParseGradientFill will expect that called ParseColor returns
  // XLColor value.
  .AddComplexTypeMapping("CT_Color", "XLColor")
  // Code generator will expect that method ParseGradientStop will return
  // a named tuple. In this case, the method ParseGradientStop is generated by
  // CodeGen and developer thus has to hand-code hook with following signature
  // private (double Value, XLColor Color) OnGradientStopParsed(XLColor color, double position)
  // that will perform construct the return value.
  .AddComplexTypeMapping("CT_GradientStop", "(double Value, XLColor Color)")

new ParserGenerator(schema, typeMap, "Demo", "_ns")
  .AddParseMethod("CT_GradientFill")
  .AddParseMethod("CT_GradientStop");
```

This is an example of generated methods in current incarantion:
```csharp
private void ParseGradientFill(string elementName)
{
    // Other attributes omitted for brevity
    var type = _reader.GetOptionalEnum<XLGradientType>("type") ?? XLGradientType.Linear;

    // Because generator knows that ParseGradientStop should return value and that it can contain a sequence here, it stores values in a list
    var stop = new List<(double Value, XLColor Color)>();
    while (_reader.TryOpen("stop", _ns))
    {
        stop.Add(ParseGradientStop("stop"));
    }
    _reader.Close(elementName, _ns);

    // Extracted valkus are supplied to the partial hook. The ParseGradientFill doesn't have a hook and thus doesn't return a value.
    OnGradientFillParsed(stop, type, degree, left, right, top, bottom);
}

private (double Value, XLColor Color) ParseGradientStop(string elementName)
{
    var position = _reader.GetDouble("position");
    _reader.Open("color", _ns);
    var color = ParseColor("color");
    _reader.Close(elementName, _ns);

    // Generated code calls hand-coded hook in a separate partial class
    return OnGradientStopParsed(color, position);
}
```

The mapping allows to use a composition for some elements. Without this feature, it would be necessary to store `GradientStop` values in some private property and use them later in the `OnGradientFillParsed`. That pattern is feasible, but hard to read.

## Hooks

Each generated `Parse*` method calls a hook method once it is completely parsed and all values of the element (attributes and mapped complex types) are passed to the hook. The hook method is generally a partial method and thus doesn't have to be implemented (unless it is a hook for mapped complex type).
