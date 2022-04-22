using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System;
using System.IO;
using System.Reflection;

namespace ClosedXML.Examples
{
    public class ImageFormats : IXLExample, IDisposable
    {
        private bool disposedValue;
        private XLWorkbook wb;

        public void Create(string filePath)
        {
            wb = new XLWorkbook();
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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    wb?.Dispose();
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}