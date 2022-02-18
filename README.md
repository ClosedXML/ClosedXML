# ClosedXML
[![Release](https://img.shields.io/badge/release-0.95.2-blue.svg)](https://github.com/ClosedXML/ClosedXML/releases/latest) [![NuGet Badge](https://buildstats.info/nuget/ClosedXML)](https://www.nuget.org/packages/ClosedXML/) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.0-red.svg)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![Build status](https://ci.appveyor.com/api/projects/status/wobbmnlbukxejjgb?svg=true)](https://ci.appveyor.com/project/ClosedXML/ClosedXML/branch/develop/artifacts)
[![.NET build and test](https://github.com/stesee/ClosedXML/actions/workflows/dotnet.yml/badge.svg)](https://github.com/stesee/ClosedXML/actions/workflows/dotnet.yml) [![NuGet Badge](https://buildstats.info/nuget/DocumentPartner.ClosedXML)](https://www.nuget.org/packages/DocumentPartner.ClosedXML/)

ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

This fork adds linux support.

[For more information see the wiki](https://github.com/closedxml/closedxml/wiki)

## Install ClosedXML via NuGet

To install ClosedXML, run the following command in the Package Manager Console

``` powershell
PM> Install-Package DocumentPartner.ClosedXML
```

## What can you do with this?

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

### Frequent answers
- ClosedXML is not thread-safe. There is no guarantee that [parallel operations](https://github.com/ClosedXML/ClosedXML/issues/1662) will work. The underlying OpenXML library is also not thread-safe.
- If you get an exception `The type initializer for 'Gdip' threw an exception.` on Linux, try [these](https://stackoverflow.com/a/67092403/179494) [solutions](https://github.com/dotnet/runtime/issues/27200#issuecomment-415327256).

## Extensions

Be sure to check out our `ClosedXML` extension projects

- <https://github.com/ClosedXML/ClosedXML.Report>
- <https://github.com/ClosedXML/ClosedXML.Extensions.AspNet>
- <https://github.com/ClosedXML/ClosedXML.Extensions.Mvc>
- <https://github.com/ClosedXML/ClosedXML.Extensions.WebApi>

## Developer guidelines

The [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) is a large and complicated beast. In order for ClosedXML, the wrapper around OpenXML, to support all the features, we rely on community contributions. Before opening an issue to request a new feature, we'd like to urge you to try to implement it yourself and log a pull request.

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits

- Project originally created by Manuel de Leon
- Current maintainer: [Francois Botha](https://github.com/igitur)
- Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
- Logo design by [@Tobaloidee](https://github.com/Tobaloidee)
