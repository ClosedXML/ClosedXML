using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;

namespace ClosedXML_Examples
{
    public class MultipleRanges
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Multiple Ranges");
            
            // using multiple string range definitions
            ws.Ranges("A1:B2", "C3:D4", "E5:F6").Style.Fill.BackgroundColor = Color.Red;

            // using a single string separated by commas
            ws.Ranges("A5:B6,E1:F2").Style.Fill.BackgroundColor = Color.Orange;

            workbook.SaveAs(filePath);
        }
    }
}
