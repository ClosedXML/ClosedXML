using System;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class CustomAutoFilter : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws;

            #region Single Column Numbers
            String singleColumnNumbers = "Single Column Numbers";
            ws = wb.Worksheets.Add(singleColumnNumbers);

            // Add a bunch of numbers to filter
            ws.Cell("A1").SetValue("Numbers")
                         .CellBelow().SetValue(2)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(5)
                         .CellBelow().SetValue(1)
                         .CellBelow().SetValue(4);

            // Add filters
            ws.RangeUsed().SetAutoFilter().Column(1).EqualTo(3).Or.GreaterThan(4);

            // Sort the filtered list
            ws.AutoFilter.Sort(1);
            #endregion

            #region Single Column Strings
            String singleColumnStrings = "Single Column Strings";
            ws = wb.Worksheets.Add(singleColumnStrings);

            // Add a bunch of strings to filter
            ws.Cell("A1").SetValue("Strings")
                         .CellBelow().SetValue("B")
                         .CellBelow().SetValue("C")
                         .CellBelow().SetValue("C")
                         .CellBelow().SetValue("E")
                         .CellBelow().SetValue("A")
                         .CellBelow().SetValue("D");

            // Add filters
            ws.RangeUsed().SetAutoFilter().Column(1).Between("B", "D");

            // Sort the filtered list
            ws.AutoFilter.Sort(1);
            #endregion

            #region Single Column Mixed
            String singleColumnMixed = "Single Column Mixed";
            ws = wb.Worksheets.Add(singleColumnMixed);

            // Add a bunch of items to filter
            ws.Cell("A1").SetValue("Mixed")
                         .CellBelow().SetValue("B")
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue("C")
                         .CellBelow().SetValue("E")
                         .CellBelow().SetValue(1)
                         .CellBelow().SetValue(4);

            // Add filters
            ws.RangeUsed().SetAutoFilter().Column(1).EqualTo(3).Or.EqualTo("C");

            // Sort the filtered list
            ws.AutoFilter.Sort(1);
            #endregion

            #region Multi Column
            String multiColumn = "Multi Column";
            ws = wb.Worksheets.Add(multiColumn);

            ws.Cell("A1").SetValue("First")
             .CellBelow().SetValue("B")
             .CellBelow().SetValue("C")
             .CellBelow().SetValue("C")
             .CellBelow().SetValue("E")
             .CellBelow().SetValue("A")
             .CellBelow().SetValue("D");

            ws.Cell("B1").SetValue("Numbers")
                         .CellBelow().SetValue(2)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(5)
                         .CellBelow().SetValue(1)
                         .CellBelow().SetValue(4);

            ws.Cell("C1").SetValue("Strings")
             .CellBelow().SetValue("B")
             .CellBelow().SetValue("C")
             .CellBelow().SetValue("C")
             .CellBelow().SetValue("E")
             .CellBelow().SetValue("A")
             .CellBelow().SetValue("D");

            // Add filters
            ws.RangeUsed().SetAutoFilter().Column(2).EqualTo(3).Or.GreaterThan(4);
            ws.RangeUsed().SetAutoFilter().Column(3).Between("B", "D");

            // Sort the filtered list
            ws.AutoFilter.Sort(3);
            #endregion

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);

                var workbook = new XLWorkbook(ms);

                #region Single Column Numbers
                workbook.Worksheet(singleColumnNumbers).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Single Column Strings
                workbook.Worksheet(singleColumnStrings).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Single Column Mixed
                workbook.Worksheet(singleColumnMixed).AutoFilter.Column(1).EqualOrGreaterThan("D");
                workbook.Worksheet(singleColumnMixed).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Multi Column
                workbook.Worksheet(multiColumn).AutoFilter.Column(3).EqualTo("E");
                workbook.Worksheet(multiColumn).AutoFilter.Sort(3, XLSortOrder.Descending);
                #endregion

                workbook.SaveAs(filePath);
                ms.Close();
            }
        }
    }
}
