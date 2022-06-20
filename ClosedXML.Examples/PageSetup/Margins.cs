using ClosedXML.Excel;

namespace ClosedXML.Examples.PageSetup
{
    public class Margins : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Margins");
            ws.PageSetup.Margins.Top = 1;
            ws.PageSetup.Margins.Bottom = 1.25;
            ws.PageSetup.Margins.Left = 0.5;
            ws.PageSetup.Margins.Right = 0.75;
            ws.PageSetup.Margins.Footer = 0.15;
            ws.PageSetup.Margins.Header = 0.30;

            ws.PageSetup.CenterHorizontally = true;
            ws.PageSetup.CenterVertically = true;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}