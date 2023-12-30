using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    [TestFixture]
    public class Top10FilterTests
    {
        [Test]
        public void Top10_filter_is_initialized_after_load()
        {
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    var autoFilter = ws.Cell("A1").InsertData(new object[]
                    {
                        "Data",
                        4,4,1,3,2,5,
                    }).SetAutoFilter();
                    autoFilter.Column(1).Top(3);
                },
                (_, ws) =>
                {
                    ws.AutoFilter.Reapply();
                    var filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                    CollectionAssert.AreEqual(new[] { true, true, false, false, false, true }, filterResult);
                });
        }
    }
}
