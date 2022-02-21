using System;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class TopBottomAutoFilter : IXLExample
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
            ws.RangeUsed().SetAutoFilter().Column(1).Top(2);

            // Sort the filtered list
            //ws.AutoFilter.Sort(1);
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
            ws.RangeUsed().SetAutoFilter().Column(2).Bottom(50, XLTopBottomType.Percent);

            // Sort the filtered list
            //ws.AutoFilter.Sort(3);
            #endregion

            using (var ms = new MemoryStream())
            {
                wb.SaveAs(ms);

                var workbook = new XLWorkbook(ms);

                #region Single Column Numbers
                //workbook.Worksheet(singleColumnNumbers).AutoFilter.Sort(1, XLSortOrder.Descending);
                #endregion

                #region Multi Column 
                //workbook.Worksheet(multiColumn).AutoFilter.Sort(3, XLSortOrder.Descending);
                #endregion

                workbook.SaveAs(filePath);
                ms.Close();
            }
        }
    }
}
