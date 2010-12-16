using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

using System.Drawing;

namespace ClosedXML_Examples.PageSetup
{
    public class TwoPages
    {
        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            foreach (var ro in Enumerable.Range(1, 100))
            {
                foreach (var co in Enumerable.Range(1, 10))
                {
                    ws.Cell(ro, co).Value = ws.Cell(ro, co).Address.ToString();
                }
            }
            ws.PageSetup.PagesWide = 1;

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
