using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    [TestFixture]
    public class DynamicFilterTests
    {
        [Test]
        public void Average_filter_is_initialized_after_load()
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    var autoFilter = ws.Cell("A1").InsertData(new object[]
                    {
                        "Data",
                        1,2,3,4,5,10, // avg. 4.16
                    }).SetAutoFilter();
                    autoFilter.Column(1).AboveAverage();
                },
                (_, ws) =>
                {
                    ws.AutoFilter.Reapply();
                    var filterResult = ws.Rows("2:7").Select(row => row.IsHidden);
                    CollectionAssert.AreEqual(new[] { true, true, true, true, false, false }, filterResult);
                });
        }
    }
}
