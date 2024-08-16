using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class PrintAreaTests
    {
        [Test]
        [TestCase("A1:B2")]
        [TestCase("A1:B2", "D3:D5")]
        public void CanLoadWorksheetWithMultiplePrintAreas(params string[] printAreaRangeAddresses)
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    foreach (var printAreaRangeAddress in printAreaRangeAddresses)
                        ws.PageSetup.PrintAreas.Add(printAreaRangeAddress);
                },
                (_, ws) =>
                {
                    var actualPrintAddresses = ws.PageSetup.PrintAreas.Select(pa => pa.RangeAddress.ToStringRelative());
                    CollectionAssert.AreEqual(printAreaRangeAddresses, actualPrintAddresses);
                });
        }
    }
}
