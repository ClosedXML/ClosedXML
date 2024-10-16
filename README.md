![ClosedXML](https://github.com/ClosedXML/ClosedXML/blob/develop/resources/logo/readme.png)

[![Release](https://img.shields.io/badge/release-0.95.4-blue.svg)](https://github.com/ClosedXML/ClosedXML/releases/latest) [![NuGet version (ClosedXML)](https://img.shields.io/nuget/v/ClosedXML.svg?style=flat)](https://www.nuget.org/packages/ClosedXML/) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.0-red.svg)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![Build status](https://ci.appveyor.com/api/projects/status/wobbmnlbukxejjgb?svg=true)](https://ci.appveyor.com/project/ClosedXML/ClosedXML/branch/develop/artifacts)
[![Open Source Helpers](https://www.codetriage.com/closedxml/closedxml/badges/users.svg)](https://www.codetriage.com/closedxml/closedxml)

[💾 Download unstable CI build](https://ci.appveyor.com/project/ClosedXML/ClosedXML/branch/develop/artifacts)

ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

For more information see [the documentation](https://closedxml.readthedocs.io/) or [the wiki](https://github.com/closedxml/closedxml/wiki).

### Release notes & migration guide

The public API is still not stable and it is a very good idea to **read release notes** and **migration guide** before each update.
* [Release notes for 0.100](https://github.com/ClosedXML/ClosedXML/releases/tag/0.100.0)
* [Migration guide for 0.100](https://closedxml.readthedocs.io/en/latest/migrations/migrate-to-0.100.html)
* [Release notes for 0.97](https://github.com/ClosedXML/ClosedXML/releases/tag/0.97.0)

### Performance

Performance matters mostly for large files. For small files, few ms here or there doesn't matter. The presented data are from generally develop branch (currently [0.103-beta](https://github.com/ClosedXML/ClosedXML/commit/5f7c0d9461352a6a468e5299bfef6eaf82bf37da)).

<details>
  <summary>Runtime details</summary>
```
BenchmarkDotNet v0.13.8, Windows 11 (10.0.22621.2283/22H2/2022Update/SunValley2)
AMD Ryzen 5 5500U with Radeon Graphics, 1 CPU, 12 logical and 6 physical cores
.NET SDK 8.0.100-rc.1.23463.5
  [Host] : .NET 7.0.11 (7.0.1123.42427), X64 RyuJIT AVX2
```
</details>

#### Save

| Description                  |     Rows  |           Columns      | Time/Memory to insert data | Save workbook | Total time/memory | 
|------------------------------|-----------|------------------------|----------------------------|------------------------------|---|
| Mixed (text/number) workbook.<br>[Gist](https://gist.github.com/jahav/bdc5fe3c90f25544ca6ae1394bbe3561) |   250 000 | 15 | 1.619 sec / 117 MiB |  6.343 sec |  7.962 sec /  477 MiB |
| Text only workbook. [Gist](https://gist.github.com/jahav/257bb2ffd5ab7adfded7e669290d8151)              | 1 000 000 | 10 | 6.302 sec / 402 MiB  | 17.134 sec | 23.436 sec / 1880 MiB |

#### Load

| Description                                                        |  Rows     | Columns | Time to load data | Used memory |
|--------------------------------------------------------------------|-----------|---------|-------------------|-------------|
| Load mixed workbook (10 text/5 number columns). |   250 000 |      15 |        15.648 sec |     236 MiB |
| Text only workbook.                                                | 1 000 000 |      10 |       49.046 sec  |     801 MiB |

Load tests used files created during save test.

### Frequent answers
- If you get an exception `Unable to find font font name or fallback font fallback font name. Install missing fonts or specify a different fallback font through ‘LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine(“Fallback font name”)’`, see help page about [missing fonts](https://closedxml.readthedocs.io/en/latest/tips/missing-font.html).
- ClosedXML is not thread-safe. There is no guarantee that [parallel operations](https://github.com/ClosedXML/ClosedXML/issues/1662) will work. The underlying OpenXML library is also not thread-safe.
- If you get an exception `The type initializer for 'Gdip' threw an exception.` on Linux, you have to upgrade to 0.97+.

### Install ClosedXML via NuGet

If you want to include ClosedXML in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/ClosedXML)

To install ClosedXML, run the following command in the Package Manager Console

```
PM> Install-Package ClosedXML
```

### What can you do with this?

ClosedXML allows you to create Excel files without the Excel application. The typical example is creating Excel reports on a web server.

**Example:**
```c#
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}
```

### Extensions
Be sure to check out our `ClosedXML` extension projects
- https://github.com/ClosedXML/ClosedXML.Report
- https://github.com/ClosedXML/ClosedXML.Extensions.AspNet
- https://github.com/ClosedXML/ClosedXML.Extensions.Mvc
- https://github.com/ClosedXML/ClosedXML.Extensions.WebApi

## Developer guidelines
The [OpenXML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/) is a large and complicated beast. In order for ClosedXML, the wrapper around OpenXML, to support all the features, we rely on community contributions. Before opening an issue to request a new feature, we'd like to urge you to try to implement it yourself and log a pull request.

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits
* Project originally created by Manuel de Leon
* Current maintainer: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
