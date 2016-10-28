
# ClosedXML
[![Build status](https://ci.appveyor.com/api/projects/status/wobbmnlbukxejjgb?svg=true)](https://ci.appveyor.com/project/Pyropace/closedxml)

ClosedXML makes it easier for developers to create Excel 2007/2010/2013 files. It provides a nice object oriented way to manipulate the files (similar to VBA) without dealing with the hassles of XML Documents. It can be used by any .NET language like C# and Visual Basic (VB).

### Install ClosedXML via NuGet

If you want to include ClosedXML in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/ClosedXML)

To install ClosedXML, run the following command in the Package Manager Console

```
PM> Install-Package ClosedXML
```

### What can you do with this?

ClosedXML allows you to create Excel 2007/2010/2013 files without the Excel application. The typical example is creating Excel reports on a web server.

If you've ever used the Microsoft Open XML Format SDK you know just how much code you have to write to get the same results as the following 4 lines of code.

```c#
var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Sample Sheet");
worksheet.Cell("A1").Value = "Hello World!";
workbook.SaveAs("HelloWorld.xlsx");
```
