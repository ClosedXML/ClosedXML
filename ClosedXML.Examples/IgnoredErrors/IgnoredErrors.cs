using System;
using ClosedXML.Excel;

namespace ClosedXML.Examples.IgnoredErrors
{
    public class IgnoredErrors : IXLExample
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();

            var ws1 = workbook.Worksheets.Add("IgnoredErrors1");
            ws1.Row(1).Cell(1).SetValue("11");
            ws1.IgnoredErrors.Add(XLIgnoredErrorType.NumberAsText, ws1.Range(1, 1, 1, 1));

            var ws2 = ws1.CopyTo("IgnoredErrors2");
            ws2.Row(1).Cell(2).SetValue("12");
            ws2.Row(2).Cell(1).SetValue("21");
            ws2.Row(2).Cell(2).SetValue("22");
            ws2.IgnoredErrors.Add(XLIgnoredErrorType.NumberAsText, ws2.Range(2, 2, 2, 2));

            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
