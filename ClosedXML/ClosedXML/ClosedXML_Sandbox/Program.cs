using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
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

            var sheet = wb.Worksheets.Add("orderlines");

            int row = 1;

            int column = 1;
            sheet.Cell(row, column++).Value = "Date";
            sheet.Cell(row, column++).Value = "Quantity";
            sheet.Cell(row, column++).Value = "Category";
            sheet.Cell(row, column++).Value = "Item";
            sheet.Cell(row, column++).Value = "Unit price";
            sheet.Cell(row, column++).Value = "Total price";

            // Sample data row
            row++;
            column = 1;
            sheet.Cell(row, column++).Value = new DateTime(2014, 6, 21);
            sheet.Cell(row, column++).Value = 1;
            sheet.Cell(row, column++).Value = "Widgets";
            sheet.Cell(row, column++).Value = "Pro widget";
            sheet.Cell(row, column++).Value = "1.23";
            sheet.Cell(row, column++).Value = "1.23";

            var dataRange = sheet.RangeUsed();

            // Add a new sheet for our pivot table
            var pivotTableSheet = wb.Worksheets.Add("PivotTable");

            // Create the pivot table, using the data from the "PastrySalesData" table
            var pt = pivotTableSheet.PivotTables.AddNew("PivotTable", pivotTableSheet.Cell(1, 1), dataRange);

            // The rows in our pivot table will be the names of the categories
            pt.RowLabels.Add("Item");
            pt.RowLabels.Add("Category");
            

            pt.Values.Add("Total price");
            
            wb.SaveAs(@"c:\temp\saved3.xlsx");
            Console.WriteLine("Done");
            //Console.ReadLine();
        }
    }
}
