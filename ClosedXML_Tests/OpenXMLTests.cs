using System.IO;
using ClosedXML_Examples;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class OpenXMLTests
    {
#if !APPVEYOR
        [Test]
        public static void SetPackagePropertiesEntryToNullWithOpenXml()
        {
            // Will fail until https://github.com/OfficeDev/Open-XML-SDK/issues/235 is fixed. 
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
#endif
    }
}
