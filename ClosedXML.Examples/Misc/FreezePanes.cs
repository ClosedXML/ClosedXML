using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class FreezePanes : IXLExample
    {
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            // Freeze rows and columns in one shot
            var ws1 = wb.AddWorksheet("Freeze1");
            ws1.Cell(5, 5).SetActive();
            ws1.SheetView.Freeze(3, 3);

            // You can also be more specific on what you want to freeze
            // For example:
            var ws2 = wb.AddWorksheet("FreezeRows");
            ws2.Cell(5, 5).SetActive();
            ws2.SheetView.FreezeRows(3);

            var ws3 = wb.AddWorksheet("FreezeColumns");
            ws3.Cell(5, 5).SetActive();
            ws3.SheetView.FreezeColumns(3);

            var wsSplit = wb.AddWorksheet("Split View");
            wsSplit.Cell(2, 2).SetActive();
            wsSplit.SheetView.SplitRow = 3;
            wsSplit.SheetView.SplitColumn = 3;

            wb.SaveAs(filePath);
        }
    }
}
