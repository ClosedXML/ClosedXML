# ClosedXML

[![.NET build and test](https://github.com/stesee/ClosedXML/actions/workflows/dotnet.yml/badge.svg)](https://github.com/stesee/ClosedXML/actions/workflows/dotnet.yml) [![NuGet Badge](https://buildstats.info/nuget/DocumentPartner.ClosedXML)](https://www.nuget.org/packages/DocumentPartner.ClosedXML/)

ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

This fork of [ClosedXML](https://www.nuget.org/packages/ClosedXML/) adds Linux and MacOs support.

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

## Developer guidelines

The [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) is a large and complicated beast. In order for ClosedXML, the wrapper around OpenXML, to support all the features, we rely on community contributions. Before opening an issue to request a new feature, we'd like to urge you to try to implement it yourself and log a pull request.

Please read the [full developer guidelines](CONTRIBUTING.md).
