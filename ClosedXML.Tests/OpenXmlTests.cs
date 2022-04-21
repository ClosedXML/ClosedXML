using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class OpenXmlTests
    {
        [Test]
        public static void SetPackagePropertiesEntryToNullWithOpenXml()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);

                using (var document = SpreadsheetDocument.Open(ms, true))
                {
                    document.PackageProperties.Creator = null;
                }
            }
        }
    }
}
