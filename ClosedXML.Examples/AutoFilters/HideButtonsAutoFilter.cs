using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class HideButtonsAutoFilter : IXLExample
    {
        public void Create(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("All buttons visible");

                for (int i = 1; i <= 3; i++)
                {
                    ws1.Column(i).Width = 20;

                    ws1.Cell(1, i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws1.Cell(1, i).Style.Font.Bold = true;
                    ws1.Cell(1, i).Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                ws1.Cell(1, 1).SetValue("Column 1")
                    .CellBelow().SetValue(1)
                    .CellBelow().SetValue(2)
                    .CellBelow().SetValue(3);
                ws1.Cell(1, 2).SetValue("Column 2")
                    .CellBelow().SetValue(4)
                    .CellBelow().SetValue(5)
                    .CellBelow().SetValue(6);

                // merged header... button is visible
                ws1.Cell(1, 1).SetValue("Column 1 & 2");
                ws1.Range(1, 1, 1, 2).Merge();

                ws1.Cell(1, 3).SetValue("Column 3")
                    .CellBelow().SetValue(7)
                    .CellBelow().SetValue(8)
                    .CellBelow().SetValue(9);

                ws1.RangeUsed().SetAutoFilter();

                var ws2 = wb.Worksheets.Add("Some buttons hidden");

                for (int i = 1; i <= 3; i++)
                {
                    ws2.Column(i).Width = 20;

                    ws2.Cell(1, i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws2.Cell(1, i).Style.Font.Bold = true;
                    ws2.Cell(1, i).Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                ws2.Cell(1, 1).SetValue("Column 1")
                    .CellBelow().SetValue(1)
                    .CellBelow().SetValue(2)
                    .CellBelow().SetValue(3);
                ws2.Cell(1, 2).SetValue("Column 2")
                    .CellBelow().SetValue(4)
                    .CellBelow().SetValue(5)
                    .CellBelow().SetValue(6);

                // merged header... button is visible
                ws2.Cell(1, 1).SetValue("Column 1 & 2");
                ws2.Range(1, 1, 1, 2).Merge();

                ws2.Cell(1, 3).SetValue("Column 3")
                    .CellBelow().SetValue(7)
                    .CellBelow().SetValue(8)
                    .CellBelow().SetValue(9);

                var autoFilter = ws2.RangeUsed().SetAutoFilter();
                // hide the button for the merged header
                autoFilter.Column(1).HideButton = true;

                wb.SaveAs(filePath);
            }
        }
    }
}
