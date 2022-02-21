using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.IO;
using System.Reflection;

namespace ClosedXML.Examples
{
    public class ImageFormats : IXLExample
    {
        public void Create(string filePath)
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws;

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Examples.Resources.ImageHandling.jpg"))
            {
                #region Jpeg

                ws = wb.Worksheets.Add("Jpg");
                ws.AddPicture(fs, XLPictureFormat.Jpeg, "JpegImage")
                    .MoveTo(ws.Cell(1, 1));

                #endregion Jpeg
            }

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Examples.Resources.ImageHandling.png"))
            {
                #region Png

                ws = wb.Worksheets.Add("Png");
                ws.AddPicture(fs, XLPictureFormat.Png, "PngImage")
                    .MoveTo(ws.Cell(1, 1));

                #endregion Png

                wb.SaveAs(filePath);
            }
        }
    }
}
