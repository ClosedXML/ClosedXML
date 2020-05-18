using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples
{
    public class ColouredAutoFilter : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws;

            #region Single Coloured Mixed

            String singleColouredNumbers = "Single Column Mixed";
            ws = wb.Worksheets.Add(singleColouredNumbers);

            Int32 ro = 1;
            Int32 co = 1;
            // Add a bunch of numbers to filter with filled color
            ws.Cell("A1").SetValue("Mixed")
                         .CellBelow().SetValue(1).Style.Fill.SetBackgroundColor(XLColor.AirForceBlue);
            ro++;
            ws.Cell(ro, co).SetValue("A").Style.Fill
                .SetBackgroundColor(XLColor.AirForceBlue);

            ro++;
            ws.Cell(ro, co).SetValue("B").Style.Fill.SetBackgroundColor(XLColor.Red);
            ro++;
            ws.Cell(ro, co).SetValue(1).Style.Fill.SetBackgroundColor(XLColor.Red);
            ro++;
            ws.Cell(ro, co).SetValue(2).Style.Fill.SetBackgroundColor(XLColor.AirForceBlue);
            ro++;
            ws.Cell(ro, co).SetValue("C").Style.Fill.SetBackgroundColor(XLColor.Purple);
            ro++;
            ws.Cell(ro, co).SetValue(3).Style.Fill.SetBackgroundColor(XLColor.Red);

            // Add filters
            ws.RangeUsed().SetAutoFilter().AddColorFilter(1, XLColor.Red);

            #endregion Single Coloured Mixed

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }
    }
}
