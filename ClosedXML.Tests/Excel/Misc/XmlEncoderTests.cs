using ClosedXML.Utils;
using NUnit.Framework;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class XmlEncoderTest
    {
        [Test]
        public void TestControlChars()
        {
            Assert.AreEqual("_x0001_ _x0002_ _x0003_ _x0004_", XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004"));
            Assert.AreEqual("_x0005_ _x0006_ _x0007_ _x0008_", XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008"));
        }

        [Test]
        public void AstralUnicodeCharsAreWrittenWithoutOpenXmlEncoding()
        {
            using var sr = new StreamReader(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Unicode\let_it_go_in_emoji.txt")));
            var surrogateEmoji = sr.ReadToEnd();

            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet();

                var cell = ws.FirstCell();
                cell.Value = "This emoji version of Let It Go from Frozen:";
                cell.CellBelow().Value = surrogateEmoji;

                return wb;
            }, @"Other\Unicode\let_it_go_in_emoji-outputfile.xlsx");
        }
    }
}
