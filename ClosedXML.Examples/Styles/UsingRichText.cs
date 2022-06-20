using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class UsingRichText : IXLExample
    {
        #region Variables

        // Public

        // Private

        #endregion Variables

        #region Properties

        // Public

        // Private

        // Override

        #endregion Properties

        #region Constructors

        // Public

        // Private

        #endregion Constructors

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Rich Text");

            // Let's start with a plain text and then decorate it...
            var cell1 = ws.Cell(1, 1).SetValue("The show must go on...");

            // We want everything in blue except the word show 
            // (which we want in red and with Courier Font)
            cell1.Style.Font.FontColor = XLColor.Blue; // Set the color for the entire cell
            cell1.GetRichText().Substring(4, 4)
                .SetFontColor(XLColor.Red)
                .SetFontName("Courier"); // Set the color and font for the word "show"

            // On the next example we'll start with an empty cell and add the rich text
            var cell = ws.Cell(3, 1);

            // Add the text parts
            cell.GetRichText().AddText("Hello").SetFontColor(XLColor.Red);
            cell.GetRichText().AddText(" BIG ").SetFontColor(XLColor.Blue).SetBold();
            cell.GetRichText().AddText("World").SetFontColor(XLColor.Red);

            // Here we're showing that even though we added three pieces of text
            // you can treat then like a single one.
            cell.GetRichText().Substring(4, 7).SetUnderline();

            // Right now cell.RichText has the following 5 strings:
            // 
            // "Hell"  -> Red
            // "o"     -> Red, Underlined
            // " BIG " -> Blue, Underlined, Bold
            // "W"     -> Red, Underlined
            // "orld"  -> Red

            // Of course you can loop through each piece of text and check its properties
            foreach (var richText in cell.GetRichText())
            {
                if (richText.Bold)
                {
                    ws.Cell(3, 2).Value = string.Format("\"{0}\" is Bold.", richText.Text);
                }
            }

            // Now we'll build a cell with rich text, and some other styles 
            cell = ws.Cell(5, 1);

            // Add the text parts
            cell.GetRichText().AddText("Some").SetFontColor(XLColor.Green);
            cell.GetRichText().AddText(" rich text ").SetFontColor(XLColor.Blue).SetBold();
            cell.GetRichText().AddText("with a gray background").SetItalic();

            cell.Style.Fill.SetBackgroundColor(XLColor.Gray);

            ws.Cell(5, 2).Value = cell.GetRichText(); // Should copy only rich text, but not background

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override

        #endregion Methods
    }
}