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
