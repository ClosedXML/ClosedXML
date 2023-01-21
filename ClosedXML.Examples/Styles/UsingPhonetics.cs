using System;
using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class UsingPhonetics : IXLExample
    {
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Using Phonetics");

            var cell = ws.Cell(1, 1);

            // Phonetics are implemented as part of the Rich Text functionality. For more information see [Using Rich Text]
            // First we add the text.
            cell.GetRichText().AddText("みんなさんはお元気ですか。").SetFontSize(16);

            // And then we add the phonetics
            cell.GetRichText().Phonetics.SetFontSize(8);
            cell.GetRichText().Phonetics.Add("げん", 7, 8);
            cell.GetRichText().Phonetics.Add("き", 8, 9);

            // Must set flag to actually display furigana
            cell.ShowPhonetic = true;

            wb.SaveAs(filePath);
        }
    }
}
