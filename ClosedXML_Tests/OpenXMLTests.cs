using System.IO;
using ClosedXML_Examples;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class OpenXMLTests
    {
        [Test]        
        public static void SetPackagePropertiesEntryToNullWithOpenXml () 
        {            
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);

                using (var document = SpreadsheetDocument.Open (ms, true)) 
                {
                    document.PackageProperties.Creator = null;
                }
            }
        }
    }
}
