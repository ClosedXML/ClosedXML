using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables
{
    [TestFixture]
    public class XLPivotTableStyleTests
    {
        [Test]
        public void PivotSubtotalsStylesLoadingTest()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\Styles\subtotals-different-styles-input.xlsx")))
                TestHelper.CreateAndCompare(() =>
                {
                    return new XLWorkbook(stream);
                }, @"Other\PivotTableReferenceFiles\Styles\subtotals-different-styles-input.xlsx");
        }
    }
}
