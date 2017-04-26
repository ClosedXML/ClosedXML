using System;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace ClosedXML_Examples
{
    public class ImageFormats : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            XLPicture pic;
            IXLWorksheet ws;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.jpg"))
            {
                #region Jpeg
                ws = wb.Worksheets.Add("Jpg");
                pic = new XLPicture()
                {
                    IsAbsolute = false,
                    ImageStream = fs,
                    Name = "JpegImage",
                    Type = "jpeg",
                    OffsetX = 0,
                    OffsetY = 0
                };

                pic.AddMarker(new XLMarker
                {
                    ColumnId = 0,
                    RowId = 0
                });

                ws.AddPicture(pic);
                #endregion
            }

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
            {
                #region Png
                ws = wb.Worksheets.Add("Png");
                pic = new XLPicture()
                {
                    IsAbsolute = false,
                    ImageStream = fs,
                    Name = "PngImage",
                    Type = "png",
                    OffsetX = 0,
                    OffsetY = 0
                };

                pic.AddMarker(new XLMarker
                {
                    ColumnId = 0,
                    RowId = 0
                });

                ws.AddPicture(pic);
                #endregion

                wb.SaveAs(filePath);
            }
        }
    }
}
