using ClosedXML.Excel;
using System;

namespace ClosedXML.Examples
{
    public class DateTimeGroupAutoFilter : IXLExample
    {
        public void Create(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws;

                #region Single Column Dates

                String singleColumnDates = "Single Column Dates";
                ws = wb.Worksheets.Add(singleColumnDates);

                // Add a bunch of numbers to filter
                ws.Cell("A1").SetValue("Dates")
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(2))
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(3))
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(3))
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(5))
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(1))
                             .CellBelow().SetValue(new DateTime(2018, 1, 1).AddDays(4));

                ws.Column(1).Style.NumberFormat.Format = "d MMMM yyyy";

                // Add filters
                ws.RangeUsed().SetAutoFilter().Column(1).AddDateGroupFilter(new DateTime(2018, 1, 1).AddDays(3), XLDateTimeGrouping.Day);

                // Sort the filtered list
                ws.AutoFilter.Sort(1);

                #endregion Single Column Dates

                ws.Columns().AdjustToContents();
                wb.SaveAs(filePath);
            }
        }
    }
}
