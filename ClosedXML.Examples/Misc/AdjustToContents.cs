using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class AdjustToContents : IXLExample
    {
        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Adjust To Contents");

            // Set some values with different font sizes
            ws.Cell(1, 1).Value = "Tall Row";
            ws.Cell(1, 1).Style.Font.FontSize = 30;
            ws.Cell(2, 1).Value = "Very Wide Column";
            ws.Cell(2, 1).Style.Font.FontSize = 20;

            // Adjust column width
            ws.Column(1).AdjustToContents();

            // Adjust row heights
            ws.Rows(1, 2).AdjustToContents();

            // You can also adjust all rows/columns in one shot
            // ws.Rows().AdjustToContents();
            // ws.Columns().AdjustToContents();

            // We'll now select which cells should be used for calculating the
            // column widths (same method applies for row heights)

            // Set the values
            ws.Cell(4, 2).Value = "Width ignored because calling column.AdjustToContents(5, 7)";
            ws.Cell(5, 2).Value = "Short text";
            ws.Cell(6, 2).Value = "Width ignored because it's part of a merge";
            ws.Range(6, 2, 6, 4).Merge();
            ws.Cell(7, 2).Value = "Width should adjust to this cell";
            ws.Cell(8, 2).Value = "Width ignored because calling column.AdjustToContents(5, 7)";

            // Adjust column widths only taking into account rows 5-7
            // (merged cells will be ignored)
            ws.Column(2).AdjustToContents(5, 7);

            // You can also specify the starting row to start calculating the widths:
            // e.g. ws.Column(3).AdjustToContents(9);

            var ws2 = wb.Worksheets.Add("Adjust Widths");
            ws2.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (var co = 0; co < 90; co += 5)
            {
                ws2.Cell(1, (co / 5) + 2).SetValue("Text to adjust - " + co).Style.Alignment.TextRotation = co;
            }

            ws2.Columns().AdjustToContents();

            var ws4 = wb.Worksheets.Add("Adjust Widths 2");
            ws4.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (var co = 0; co < 90; co += 5)
            {
                var c = ws4.Cell(1, (co / 5) + 2);

                c.GetRichText().AddText("Text to adjust - " + co).SetBold();
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("World!").SetBold().SetFontColor(XLColor.Blue).SetFontSize(25);
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("Hello Cruel and unsusual world").SetBold().SetFontSize(20);
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("Hello").SetBold();
                c.Style.Alignment.SetTextRotation(co);
            }
            ws4.Columns().AdjustToContents();

            var ws3 = wb.Worksheets.Add("Adjust Heights");
            ws3.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (var ro = 0; ro < 90; ro += 5)
            {
                ws3.Cell((ro / 5) + 2, 1).SetValue("Text to adjust - " + ro).Style.Alignment.TextRotation = ro;
            }

            ws3.Rows().AdjustToContents();

            var ws5 = wb.Worksheets.Add("Adjust Heights 2");
            ws5.Cell(1, 1).SetValue("Text to adjust - 255").Style.Alignment.TextRotation = 255;
            for (var ro = 0; ro < 90; ro += 5)
            {
                var c = ws5.Cell((ro / 5) + 2, 1);
                c.GetRichText().AddText("Text to adjust - " + ro).SetBold();
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("World!").SetBold().SetFontColor(XLColor.Blue).SetFontSize(10);
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("Hello Cruel and unsusual world").SetBold().SetFontSize(15);
                c.GetRichText().AddText(XLConstants.NewLine);
                c.GetRichText().AddText("Hello").SetBold();
                c.Style.Alignment.SetTextRotation(ro);
            }

            ws5.Rows().AdjustToContents();

            var ws6 = wb.Worksheets.Add("Absurdly wide column");
            ws6.Cell("A1").Value = "Some string";

            // This column's width should be capped at 255
            ws6.Cell("B1").Value = @"Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";

            ws6.Columns().AdjustToContents();

            wb.SaveAs(filePath, true);
        }
    }
}
