using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    [TestFixture]
    public class RegularFilterTests
    {
        [Test]
        public void DateTimeGrouping_and_regular_values_can_be_used_together()
        {
            // OpenXML SDK validator considers filter and dateTimeGroup filter elements together to
            // be an error, but it isn't (XSD allows and Excel reads). Therefore, disable
            // validation for the test.
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    var autoFilter = ws.Cell("A1").InsertData(new object[]
                    {
                        "Data",
                        1, 2,
                        new DateTime(2015, 7, 25),
                        new DateTime(2015, 8, 25),
                    }).SetAutoFilter();
                    autoFilter.Column(1)
                        .AddFilter(1)
                        .AddDateGroupFilter(new DateTime(2015, 8, 1), XLDateTimeGrouping.Month);
                },
                (_, ws) =>
                {
                    ws.AutoFilter.Reapply();
                    var dataVisibility = ws.Rows("2:5").Select(row => !row.IsHidden);
                    CollectionAssert.AreEqual(new[] { true, false, false, true }, dataVisibility);
                }, false);
        }
    }
}
