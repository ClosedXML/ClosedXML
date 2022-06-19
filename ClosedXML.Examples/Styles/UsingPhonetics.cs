using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class UsingPhonetics : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Using Phonetics");

            var cell = ws.Cell(1, 1);

            // Phonetics are implemented as part of the Rich Text functionality. For more information see [Using Rich Text]
            // First we add the text.
            cell.GetRichText().AddText("みんなさんはお元気ですか。").SetFontSize(16);

            // And then we add the phonetics
            cell.GetRichText().Phonetics.SetFontSize(8);
            cell.GetRichText().Phonetics.Add("げん", 7, 8);
            cell.GetRichText().Phonetics.Add("き", 8, 9);

            //TODO: I'm looking for someone who understands Japanese to confirm the validity of the above code.

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}