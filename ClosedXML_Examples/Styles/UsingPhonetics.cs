using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples.Styles
{
    public class UsingPhonetics : IXLExample
    {
    

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Using Phonetics");

            var cell = ws.Cell(1, 1);

            // Phonetics are implemented as part of the Rich Text functionality. For more information see [Using Rich Text]
            // First we add the text.
            cell.RichText.AddText("みんなさんはお元気ですか。").SetFontSize(16);

            // And then we add the phonetics
            cell.RichText.Phonetics.SetFontSize(8);
            cell.RichText.Phonetics.Add("げん", 7, 1);
            cell.RichText.Phonetics.Add("き", 8, 1);

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
