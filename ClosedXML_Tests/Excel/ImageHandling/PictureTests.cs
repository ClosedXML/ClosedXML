using System;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NUnit.Framework;

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
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker
            {
                ColumnId = 1,
                RowId = 1,
                ColumnOffset = 100,
                RowOffset = 0
            };

            using (Stream fs = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML_Tests.Resource.Images.ImageHandling.png"))
            {
                XLPicture pic = new XLPicture
                {
                    IsAbsolute = false,
                    ImageStream = fs,
                    Name = "Image1",
                    Type = "png",
                    OffsetX = 200,
                    OffsetY = 155
                };

                fs.Position = 0;
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(fs);

                // Get these values manually as they vary from machine to machine
                float horizontalRes = bitmap.HorizontalResolution;
                float verticalRes = bitmap.VerticalResolution;

                pic.AddMarker(firstMarker);

                Assert.AreEqual(false, pic.IsAbsolute);
                Assert.AreEqual("Image1", pic.Name);
                Assert.AreEqual("png", pic.Type);
                Assert.AreEqual(1, pic.GetMarkers().Count);
                Assert.AreNotEqual(null, new XLPicture().GetMarkers());
                Assert.AreEqual((long)(914400 * 252 / horizontalRes), pic.Width);
                Assert.AreEqual((long)(914400 * 152 / verticalRes), pic.Height);
                Assert.AreEqual(252, pic.RawWidth);
                Assert.AreEqual(152, pic.RawHeight);
                Assert.AreEqual((long)(914400 * 200 / horizontalRes), pic.OffsetX);
                Assert.AreEqual((long)(914400 * 155 / verticalRes), pic.OffsetY);
                Assert.AreEqual(200, pic.RawOffsetX);
                Assert.AreEqual(155, pic.RawOffsetY);
            }
        }
    }
}
