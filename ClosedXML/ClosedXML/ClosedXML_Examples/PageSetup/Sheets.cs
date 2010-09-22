using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.PageSetup
{
    public class Sheets
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws1 = workbook.Worksheets.Add("Separate PrintAreas");
            ws1.PageSetup.PrintAreas.Add(ws1.Range("A1:B2"));
            ws1.PageSetup.PrintAreas.Add(ws1.Range("D3:D5"));

            var ws2 = workbook.Worksheets.Add("Page Breaks");
            ws2.PageSetup.PrintAreas.Add(ws2.Range("A1:D5"));
            ws2.PageSetup.AddPageBreak(ws2.Row(2));
            ws2.PageSetup.AddPageBreak(ws2.Column(2));
            
            workbook.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
