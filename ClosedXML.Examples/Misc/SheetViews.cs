using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class SheetViews : IXLExample
    {
        public void Create(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws;

                ws = wb.AddWorksheet("ZoomScale");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScale = 50;

                ws = wb.AddWorksheet("ZoomScaleNormal");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScaleNormal = 70;

                ws = wb.AddWorksheet("ZoomScalePageLayoutView");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScalePageLayoutView = 85;

                ws = wb.AddWorksheet("ZoomScaleSheetLayoutView");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScaleSheetLayoutView = 120;

                ws = wb.AddWorksheet("ZoomScaleTooSmall");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScale = 5;

                ws = wb.AddWorksheet("ZoomScaleTooBig");
                ws.FirstCell().SetValue(ws.Name);
                ws.SheetView.ZoomScale = 500;

                ws = wb.AddWorksheet("TopLeftCell");
                ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;

                wb.SaveAs(filePath);
            }
        }
    }
}
