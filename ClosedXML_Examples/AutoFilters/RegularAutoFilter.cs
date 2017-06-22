using System;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class RegularAutoFilter : IXLExample
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
            ws.RangeUsed().SetAutoFilter().Column(1).AddFilter(3)
                                                    .AddFilter(1);

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
            ws.RangeUsed().SetAutoFilter().Column(1).AddFilter("C")
                                                    .AddFilter("A");

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
            ws.RangeUsed().SetAutoFilter().Column(1).AddFilter("C")
                                                    .AddFilter(1);

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
            ws.RangeUsed().SetAutoFilter().Column(2).AddFilter(3)
                                                    .AddFilter(1);

            // Sort the filtered list
            ws.AutoFilter.Sort(3);
            #endregion

            #region Table
            String tableSheetName = "Table";
            ws = wb.Worksheets.Add(tableSheetName);

            // Add a bunch of numbers to filter
            ws.Cell("A1").SetValue("Numbers")
                         .CellBelow().SetValue(2)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(3)
                         .CellBelow().SetValue(5)
                         .CellBelow().SetValue(1)
                         .CellBelow().SetValue(4);

            // Add filters
            var table = ws.RangeUsed().CreateTable();
            table.ShowTotalsRow = true;
            table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.AutoFilter.Column(1).AddFilter(3).AddFilter(4);

            table.AutoFilter.Sort(1);
            #endregion

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);

                var workbook = new XLWorkbook(ms);

                #region Single Column Numbers
                workbook.Worksheet(singleColumnNumbers).AutoFilter.Column(1).AddFilter(5);
                workbook.Worksheet(singleColumnNumbers).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Single Column Strings
                workbook.Worksheet(singleColumnStrings).AutoFilter.Column(1).AddFilter("E");
                workbook.Worksheet(singleColumnStrings).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Single Column Mixed
                workbook.Worksheet(singleColumnMixed).AutoFilter.Column(1).AddFilter("E");
                workbook.Worksheet(singleColumnMixed).AutoFilter.Column(1).AddFilter(3);
                workbook.Worksheet(singleColumnMixed).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Multi Column 
                workbook.Worksheet(multiColumn).AutoFilter.Column(3).AddFilter("C");
                workbook.Worksheet(multiColumn).AutoFilter.Sort(3, XLSortOrder.Descending);
                #endregion

                #region Table
                workbook.Worksheet(tableSheetName).Table(0).AutoFilter.Column(1).AddFilter(5);
                workbook.Worksheet(tableSheetName).Table(0).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                workbook.SaveAs(filePath);
                ms.Close();
            }
        }
    }
}
