using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples.Misc
{
    public class AutoFilter : IXLExample
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.RangeUsed().SetAutoFilter();
            
            // Your can turn off the autofilter in three ways:
            // 1) worksheet.AutoFilterRange.SetAutoFilter(false)
            // 2) worksheet.AutoFilterRange = null
            // 3) Pick any range in the worksheet and call range.SetAutoFilter(false);

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
