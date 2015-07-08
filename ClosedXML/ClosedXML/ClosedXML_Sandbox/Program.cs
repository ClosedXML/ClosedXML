using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
namespace ClosedXML_Sandbox
{
  class Program
  {
    private static void Main(string[] args)
    {
      var wb = new XLWorkbook();
      var worksheet = wb.Worksheets.Add("Sample Sheet");
      worksheet.Cell("A1").Value = "Hello World!";
      FileStream fIn = new FileStream("20150625_083814.jpg", FileMode.Open);

      XLPicture pic = new XLPicture
      {
        NoChangeAspect = true,
        NoMove = true,
        NoResize = true,
        ImageStream = fIn,
        Name = "Test Image"
      };
      XLMarker fMark = new XLMarker
      {
        ColumnId = 2,
        RowId = 2
      };
      pic.AddMarker(fMark);

      worksheet.AddPicture(pic);
      
      wb.SaveAs(@"c:\temp\saved8.xlsx");
      Console.WriteLine("Done");
      //PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);

      Console.WriteLine("Press any key to continue");
      Console.ReadKey();
    }
  }
}
