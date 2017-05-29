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
            IXLPicture picture;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
            {
                ws = wb.Worksheets.Add("Images");

                #region AbsoluteAnchor

                ws.AddPicture("Image10", fs, XLPictureFormat.Png)
                    .SetAbsolute()
                    .AtPosition(220, 150);

                #endregion AbsoluteAnchor

                #region OneCellAnchor

                fs.Position = 0;
                ws.AddPicture("Image11", fs, XLPictureFormat.Png)
                    .SetAbsolute(false)
                    .AtPosition(0, 0)
                    .WithMarker(new XLMarker
                    {
                        ColumnId = 1,
                        RowId = 1
                    });

                #endregion OneCellAnchor

                ws = wb.Worksheets.Add("MoreImages");

                #region TwoCellAnchor

                fs.Position = 0;
                picture = ws.AddPicture("Image20", fs, XLPictureFormat.Png)
                    .SetAbsolute(false)
                    .AtPosition(0, 0);

                picture.Markers.Add(new XLMarker
                {
                    ColumnId = 5,
                    RowId = 6
                });

                picture.Markers.Add(new XLMarker
                {
                    ColumnId = 7,
                    RowId = 9
                });

                #endregion TwoCellAnchor

                wb.SaveAs(filePath);
            }
        }
    }
}
