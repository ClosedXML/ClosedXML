using System;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class Hyperlinks : IXLExample
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
            var ws = wb.Worksheets.Add("Hyperlinks");
            wb.Worksheets.Add("Second Sheet");

            Int32 ro = 0;

            // You can create a link with pretty much anything you can put on a
            // browser: http, ftp, mailto, gopher, news, nntp, etc.
            
            ws.Cell(++ro, 1).Value = "Link to a web page, no tooltip - Yahoo!";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink(@"http://www.yahoo.com"));

            ws.Cell(++ro, 1).Value = "Link to a web page, with a tooltip - Yahoo!";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink(@"http://www.yahoo.com", "Click to go to Yahoo!"));

            ws.Cell(++ro, 1).Value = "Link to a file - same folder";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink("Test.xlsx"));

            ws.Cell(++ro, 1).Value = "Link to a file - Absolute";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink(@"D:\Test.xlsx"));

            ws.Cell(++ro, 1).Value = "Link to a file - relative address";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink(@"../Test.xlsx"));

            ws.Cell(++ro, 1).Value = "Link to an address in this worksheet";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink("B1"));

            ws.Cell(++ro, 1).Value = "Link to an address in another worksheet";
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink("'Second Sheet'!A1"));

            // You can also set the properties of a hyperlink directly:

            ws.Cell(++ro, 1).Value = "Link to a range in this worksheet";
            ws.Cell(ro, 1).GetHyperlink().InternalAddress = "B1:C2";
            ws.Cell(ro, 1).GetHyperlink().Tooltip = "SquareBox";

            ws.Cell(++ro, 1).Value = "Link to an email message";
            ws.Cell(ro, 1).GetHyperlink().ExternalAddress = new Uri(@"mailto:SantaClaus@NorthPole.com?subject=Presents");

            // Deleting a hyperlink
            ws.Cell(++ro, 1).Value = "This is no longer a link";
            ws.Cell(ro, 1).GetHyperlink().InternalAddress = "A1";
            ws.Cell(ro, 1).GetHyperlink().Delete();

            // Setting a hyperlink preserves previous formatting:
            ws.Cell(++ro, 1).Value = "Odd looking link";
            ws.Cell(ro, 1).Style.Font.FontColor = XLColor.Red;
            ws.Cell(ro, 1).Style.Font.Underline = XLFontUnderlineValues.Double;
            ws.Cell(ro, 1).SetHyperlink(new XLHyperlink(ws.Range("B1:C2")));
            
            // Hyperlink via formula
            ws.Cell( ++ro, 1 ).SetValue( "Send Email" )
                .SetFormulaA1( "=HYPERLINK(\"mailto:test@test.com\", \"Send Email\")" )
                .SetHyperlink(new XLHyperlink( "mailto:test@test.com", "'Send Email'" ));

            // List all hyperlinks in a worksheet:
            var hyperlinksInWorksheet = ws.Hyperlinks;

            // List all hyperlinks in a range:
            var hyperlinksInRange = ws.Range("A1:A3").Hyperlinks;

            // Clearing a cell with a hyperlink
            ws.Cell(++ro, 1).Value = "ERROR!";
            ws.Cell(ro, 1).GetHyperlink().InternalAddress = "A1";
            ws.Cell(ro, 1).Clear();

            // Deleting a cell with a hyperlink
            ws.Cell(++ro, 1).Value = "ERROR!";
            ws.Cell(ro, 1).GetHyperlink().InternalAddress = "A1";
            ws.Cell(ro, 1).Clear();

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
