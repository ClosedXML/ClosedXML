using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;
using System;
using System.IO;
using System.Reflection;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class PictureTests
    {
        [Test]
        public void XLMarkerTests()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker
            {
                ColumnId = 1,
                RowId = 1,
                ColumnOffset = 100,
                RowOffset = 0
            };

            firstMarker.ColumnId = 10;

            Assert.AreEqual(10, firstMarker.ColumnId);
            Assert.AreEqual(1, firstMarker.RowId);
            Assert.AreEqual(100, firstMarker.ColumnOffset);
            Assert.AreEqual(0, firstMarker.RowOffset);
            Assert.AreEqual(9, firstMarker.GetZeroBasedColumn());
            Assert.AreEqual(0, firstMarker.GetZeroBasedRow());

            Assert.Throws(typeof(ArgumentOutOfRangeException), delegate { firstMarker.RowId = 0; });
            Assert.Throws(typeof(ArgumentOutOfRangeException), delegate { firstMarker.ColumnId = 0; });
            Assert.Throws(typeof(ArgumentOutOfRangeException),
                            delegate { firstMarker.RowId = XLHelper.MaxRowNumber + 1; });
            Assert.Throws(typeof(ArgumentOutOfRangeException),
                            delegate { firstMarker.ColumnId = XLHelper.MaxColumnNumber + 1; });
        }

        [Test]
        public void XLPictureTests()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                XLMarker firstMarker = new XLMarker
                {
                    ColumnId = 1,
                    RowId = 1,
                    ColumnOffset = 100,
                    RowOffset = 0
                };

                using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
                {
                    var pic = ws.AddPicture("Image1", fs, XLPictureFormat.Png)
                        .SetAbsolute(false)
                        .AtPosition(220, 155);

                    fs.Position = 0;

                    pic.Markers.Add(firstMarker);

                    Assert.AreEqual(false, pic.IsAbsolute);
                    Assert.AreEqual("Image1", pic.Name);
                    Assert.AreEqual(XLPictureFormat.Png, pic.Format);
                    Assert.AreEqual(1, pic.Markers.Count);
                    Assert.AreEqual(252, pic.Width);
                    Assert.AreEqual(152, pic.Height);
                    Assert.AreEqual(220, pic.Left);
                    Assert.AreEqual(155, pic.Top);
                }
            }
        }
    }
}
