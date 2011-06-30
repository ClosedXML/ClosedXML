using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples.Styles
{
    public class UsingRichText : IXLExample
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Constructors

        // Public



        // Private


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Rich Text");

            // Let's start with a plain text and then decorate it...
            var cell1 = ws.Cell(1, 1).SetValue("The show must go on...");

            // We want everything in blue except the word show 
            // (which we want in red and with Broadway Font)
            cell1.Style.Font.FontColor = XLColor.Blue; // Set the color for the entire cell
            cell1.RichText.Substring(4, 4)
                .SetFontColor(XLColor.Red)
                .SetFontName("Broadway"); // Set the color and font for the word "show"

            // On the next example we'll start with an empty cell and add the rich text
            var cell = ws.Cell(3, 1);

            // Add the text parts
            cell.RichText.AddText("Hello").SetFontColor(XLColor.Red);
            cell.RichText.AddText(" BIG ").SetFontColor(XLColor.Blue).SetBold();
            cell.RichText.AddText("World").SetFontColor(XLColor.Red);

            // Here we're showing that even though we added three pieces of text
            // you can treat then like a single one.
            cell.RichText.Substring(4, 7).SetUnderline();

            // Right now cell.RichText has the following 5 strings:
            // 
            // "Hell"  -> Red
            // "o"     -> Red, Underlined
            // " BIG " -> Blue, Underlined, Bold
            // "W"     -> Red, Underlined
            // "orld"  -> Red

            // Of course you can loop through each piece of text and check its properties
            foreach (var richText in cell.RichText)
            {
                if(richText.Bold)
                    ws.Cell(3, 2).Value = String.Format("\"{0}\" is Bold.", richText.Text);
            }
            
            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
