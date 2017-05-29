using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.IO;
using System.Reflection;

namespace ClosedXML_Examples
{
    public class ImageFormats : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.jpg"))
            {
                #region Jpeg

                ws = wb.Worksheets.Add("Jpg");
                ws.AddPicture("JpegImage", fs, XLPictureFormat.Jpeg)
                    .SetAbsolute(false)
                    .AtPosition(0, 0)
                    .WithMarker(new XLMarker
                    {
                        ColumnId = 1,
                        RowId = 1
                    });

                #endregion Jpeg
            }

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Examples.Resources.ImageHandling.png"))
            {
                #region Png

                ws = wb.Worksheets.Add("Png");
                ws.AddPicture("PngImage", fs, XLPictureFormat.Png)
                    .SetAbsolute(false)
                    .AtPosition(0, 0)
                    .WithMarker(new XLMarker
                    {
                        ColumnId = 1,
                        RowId = 1
                    });

                #endregion Png

                wb.SaveAs(filePath);
            }
        }
    }
}
