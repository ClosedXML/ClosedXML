using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.IO;
using System.Reflection;

namespace ClosedXML_Examples
{
    public class ImageAnchors : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
            {
                ws = wb.Worksheets.Add("Images");

                #region AbsoluteAnchor

                ws.AddPicture(fs, XLPictureFormat.Png, "Image10")
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .AtPosition(220, 150);

                #endregion AbsoluteAnchor

                #region OneCellAnchor

                fs.Position = 0;
                ws.AddPicture(fs, XLPictureFormat.Png, "Image11")
                    .WithPlacement(XLPicturePlacement.MoveAndSize)
                    .AtPosition(ws.Cell(1, 1));

                #endregion OneCellAnchor

                ws = wb.Worksheets.Add("MoreImages");

                #region TwoCellAnchor

                fs.Position = 0;
                ws.AddPicture(fs, XLPictureFormat.Png, "Image20")
                    .WithPlacement(XLPicturePlacement.MoveAndSize)
                    .AtPosition(ws.Cell(6, 5), ws.Cell(9, 7));

                #endregion TwoCellAnchor

                wb.SaveAs(filePath);
            }
        }
    }
}
