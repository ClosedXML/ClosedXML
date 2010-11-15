using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;


namespace ClosedXML_Examples
{
    public class TransposeRanges
    {
        public void Create()
        {
            var workbook = new XLWorkbook(@"C:\Excel Files\Created\BasicTable.xlsx");
            var ws = workbook.Worksheets.Worksheet(0);

            var rngTable = ws.Range("B2:F6");

            rngTable.Transpose(XLTransposeOptions.MoveCells);

            ws.Columns().AdjustToContents();

            workbook.SaveAs(@"C:\Excel Files\Created\TransposeRanges.xlsx");
        }
    }
}