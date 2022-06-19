using ClosedXML.Excel;

namespace ClosedXML.Examples.PageSetup
{
    public class Page : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws1 = workbook.Worksheets.Add("Page Setup - Page1");
            ws1.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws1.PageSetup.AdjustTo(80);
            ws1.PageSetup.PaperSize = XLPaperSize.LegalPaper;
            ws1.PageSetup.VerticalDpi = 600;
            ws1.PageSetup.HorizontalDpi = 600;

            var ws2 = workbook.Worksheets.Add("Page Setup - Page2");
            ws2.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            ws2.PageSetup.FitToPages(2, 2);     // Alternatively you can use 
                                                // ws2.PageSetup.PagesTall = #
                                                // and/or ws2.PageSetup.PagesWide = #

            ws2.PageSetup.PaperSize = XLPaperSize.LetterPaper;
            ws2.PageSetup.VerticalDpi = 600;
            ws2.PageSetup.HorizontalDpi = 600;
            ws2.PageSetup.FirstPageNumber = 5;

            workbook.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}