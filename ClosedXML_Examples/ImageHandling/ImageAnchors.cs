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
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws;

                using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
                {
                    ws = wb.Worksheets.Add("Images1");

                    #region AbsoluteAnchor

                    ws.AddPicture(fs, XLPictureFormat.Png, "Image10")
                        .MoveTo(XLMeasure.Create(220, XLMeasureUnit.Pixels), XLMeasure.Create(150, XLMeasureUnit.Pixels));

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

                using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                {
                    // Moving images around and scaling them
                    ws = wb.Worksheets.Add("Images3");

                    ws.AddPicture(fs, XLPictureFormat.Jpeg)
                        .MoveTo(ws.Cell(2, 2), XLMeasure.Create(20, XLMeasureUnit.Pixels), XLMeasure.Create(5, XLMeasureUnit.Pixels), ws.Cell(5, 5), XLMeasure.Create(30, XLMeasureUnit.Pixels), XLMeasure.Create(10, XLMeasureUnit.Pixels))
                        .MoveTo(ws.Cell(2, 2), ws.Cell(5, 5));

                    ws.AddPicture(fs, XLPictureFormat.Jpeg)
                        .MoveTo(ws.Cell(6, 2), XLMeasure.Create(2, XLMeasureUnit.Pixels), XLMeasure.Create(2, XLMeasureUnit.Pixels), ws.Cell(9, 5), XLMeasure.Create(2, XLMeasureUnit.Pixels), XLMeasure.Create(2, XLMeasureUnit.Pixels))
                        .MoveTo(ws.Cell(6, 2), XLMeasure.Create(20, XLMeasureUnit.Pixels), XLMeasure.Create(5, XLMeasureUnit.Pixels), ws.Cell(9, 5), XLMeasure.Create(30, XLMeasureUnit.Pixels), XLMeasure.Create(10, XLMeasureUnit.Pixels));

                    ws.AddPicture(fs, XLPictureFormat.Jpeg)
                        .MoveTo(ws.Cell(10, 2), XLMeasure.Create(20, XLMeasureUnit.Pixels), XLMeasure.Create(5, XLMeasureUnit.Pixels))
                        .Scale(0.2, true)
                        .MoveTo(ws.Cell(10, 1));
                }

                using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.SampleImage.jpg"))
                {
                    // Changing of placement
                    ws = wb.Worksheets.Add("Images4");

                    ws.AddPicture(fs, XLPictureFormat.Jpeg)
                        .MoveTo(XLMeasure.Create(100, XLMeasureUnit.Pixels), XLMeasure.Create(100, XLMeasureUnit.Pixels))
                        .WithPlacement(XLPicturePlacement.FreeFloating);

                    // Add and delete picture immediately
                    ws.AddPicture(fs, XLPictureFormat.Jpeg)
                        .MoveTo(XLMeasure.Create(100, XLMeasureUnit.Pixels), XLMeasure.Create(600, XLMeasureUnit.Pixels))
                        .Delete();
                }

                wb.SaveAs(filePath);
            }
        }
    }
}
