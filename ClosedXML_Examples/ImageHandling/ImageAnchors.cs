using System;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML_Examples
{
    public class ImageAnchors : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            XLPicture pic;
            IXLWorksheet ws;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
            {
                ws = wb.Worksheets.Add("Images");

                #region AbsoluteAnchor
                pic = new XLPicture()
                {
                    IsAbsolute = true,
                    ImageStream = fs,
                    Name = "Image10",
                    Type = "png",
                    OffsetX = 220,
                    OffsetY = 150
                };
                ws.AddPicture(pic);
                #endregion

                #region OneCellAnchor
                fs.Position = 0;
                pic = new XLPicture()
                {
                    IsAbsolute = false,
                    ImageStream = fs,
                    Name = "Image11",
                    Type = "png",
                    OffsetX = 0,
                    OffsetY = 0
                };

                pic.AddMarker(new XLMarker
                {
                    ColumnId = 1,
                    RowId = 1
                });

                ws.AddPicture(pic);
                #endregion

                ws = wb.Worksheets.Add("MoreImages");

                #region TwoCellAnchor
                fs.Position = 0;
                pic = new XLPicture()
                {
                    IsAbsolute = false,
                    ImageStream = fs,
                    Name = "Image20",
                    Type = "png",
                    OffsetX = 0,
                    OffsetY = 0
                };

                pic.AddMarker(new XLMarker
                {
                    ColumnId = 5,
                    RowId = 6
                });

                pic.AddMarker(new XLMarker
                {
                    ColumnId = 7,
                    RowId = 9
                });
                ws.AddPicture(pic);
                #endregion

                wb.SaveAs(filePath);
            }
        }
    }
}
