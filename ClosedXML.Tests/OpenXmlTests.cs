using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class OpenXmlTests
    {
        [Test]
        [Ignore("Workaround has been included in ClosedXML")]
        public static void SetPackagePropertiesEntryToNullWithOpenXml()
        {
            // Fixed in .NET Standard 2.1
            // See:
            //      https://github.com/OfficeDev/Open-XML-SDK/issues/235
            //      https://github.com/dotnet/corefx/issues/23795
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
