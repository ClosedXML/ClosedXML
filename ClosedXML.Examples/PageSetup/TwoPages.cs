using ClosedXML.Excel;
using System.Linq;

namespace ClosedXML.Examples.PageSetup
{
    public class TwoPages : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
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

        #endregion Methods
    }
}