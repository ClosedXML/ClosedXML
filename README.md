# ClosedXML

          we are in process of moving the project from codeplex to github
          Documentations to follow 

ClosedXML makes it easier for developers to create Excel 2007/2010 files. It provides a nice object oriented way to manipulate the files (similar to VBA) without dealing with the hassles of XML Documents. It can be used by any .NET language like C# and Visual Basic (VB).

### What can you do with this?

ClosedXML allows you to create Excel 2007/2010 files without the Excel application. The typical example is creating Excel reports on a web server.

If you've ever used the Microsoft Open XML Format SDK you know just how much code you have to write to get the same results as the following 4 lines of code.

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell("A1").Value = "Hello World!";
            workbook.SaveAs("HelloWorld.xlsx");

Something more elaborate:
