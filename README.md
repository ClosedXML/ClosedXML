
# ClosedXML
[![Build status](https://ci.appveyor.com/api/projects/status/wobbmnlbukxejjgb?svg=true)](https://ci.appveyor.com/project/Pyropace/closedxml)

ClosedXML makes it easier for developers to create Excel 2007+ (.xlsx, .xlsm, etc) files. It provides a nice object oriented way to manipulate the files (similar to VBA) without dealing with the hassles of XML Documents. It can be used by any .NET language like C# and Visual Basic (VB).

[For more information see the wiki](https://github.com/closedxml/closedxml/wiki)

### Install ClosedXML via NuGet

If you want to include ClosedXML in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/ClosedXML)

To install ClosedXML, run the following command in the Package Manager Console

```
PM> Install-Package ClosedXML
```

### What can you do with this?

ClosedXML allows you to create Excel 2007+ (.xlsx, .xlsm, etc) files without the Excel application. The typical example is creating Excel reports on a web server.

If you've ever used the Microsoft Open XML Format SDK you know just how much code you have to write to get the same results as the following 4 lines of code.

```c#
var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Sample Sheet");
worksheet.Cell("A1").Value = "Hello World!";
workbook.SaveAs("HelloWorld.xlsx");
```

### Extensions
Be sure to check out our `ClosedXML` extension projects
- https://github.com/ClosedXML/ClosedXML.Extensions.AspNet
- https://github.com/ClosedXML/ClosedXML.Extensions.Mvc

## Developer guidelines
_Full guidelines to follow later_
* Please submit pull requests that are based on the `develop` branch.
* Where possible, pull requests should include unit tests that cover as many uses cases as possible. This is especially relevant when implementing Excel functions.
* Install [NUnit 3.0 Test Adapter](https://github.com/nunit/docs/wiki/Adapter-Installation) if you want to run the test suite in Visual Studio.
* We use 4 spaces for code indentation. This is the default in Visual Studio. Don't leave any trailing white space at the end of lines or files. To make this easier, ClosedXML has an [editorconfig](http://www.editorconfig.org) configuration file. It is recommended you install editorconfig from the Visual Studio Extension Manager.
