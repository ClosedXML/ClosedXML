using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Reflection;

namespace ClosedXML.Examples
{
    public class ImageAnchors : IXLExample
    {
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws;

            using (var fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Examples.Resources.ImageHandling.png"))
            {
                ws = wb.Worksheets.Add("Images1");

                #region AbsoluteAnchor

                ws.AddPicture(fs, XLPictureFormat.Png, "Image10")
                    .MoveTo(220, 150);

                #endregion AbsoluteAnchor

                #region OneCellAnchor

                fs.Position = 0;
                ws.AddPicture(fs, XLPictureFormat.Png, "Image11")
                    .MoveTo(ws.Cell(1, 1));

                #endregion OneCellAnchor

                ws = wb.Worksheets.Add("Images2");

                #region TwoCellAnchor

                fs.Position = 0;
                ws.AddPicture(fs, XLPictureFormat.Png, "Image20")
                    .MoveTo(ws.Cell(6, 5), ws.Cell(9, 7));

                #endregion TwoCellAnchor
            }

            using (var fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
            {
                // Moving images around and scaling them
                ws = wb.Worksheets.Add("Images3");

                ws.AddPicture(fs, XLPictureFormat.Jpeg)
                    .MoveTo(ws.Cell(2, 2), 20, 5, ws.Cell(5, 5), 30, 10)
                    .MoveTo(ws.Cell(2, 2), ws.Cell(5, 5));

                ws.AddPicture(fs, XLPictureFormat.Jpeg)
                    .MoveTo(ws.Cell(6, 2), 2, 2, ws.Cell(9, 5), 2, 2)
                    .MoveTo(ws.Cell(6, 2), 20, 5, ws.Cell(9, 5), 30, 10);

                ws.AddPicture(fs, XLPictureFormat.Jpeg)
                    .MoveTo(ws.Cell(10, 2), 20, 5)
                    .Scale(0.2, true)
                    .MoveTo(ws.Cell(10, 1));
            }

            using (var fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
            {
                // Changing of placement
                ws = wb.Worksheets.Add("Images4");

                ws.AddPicture(fs, XLPictureFormat.Jpeg)
                    .MoveTo(100, 100)
                    .WithPlacement(XLPicturePlacement.FreeFloating);

                // Add and delete picture immediately
                ws.AddPicture(fs, XLPictureFormat.Jpeg)
                    .MoveTo(100, 600)
                    .Delete();
            }

            wb.SaveAs(filePath);
        }
    }
}
