using System;
using NUnit.Framework;
using System.Linq;
using ClosedXML.Excel;

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
                    var filterResult = ws.Rows("2:7").Select(row => !row.IsHidden);
                    CollectionAssert.AreEqual(new[] { false, false, false, false, true, true }, filterResult);
                });
        }

        [Test]
        public void BelowAverage_takes_values_under_avg_value()
        {
            // The average 2 is not included. 
            new AutoFilterTester(f => f.BelowAverage())
                .AddTrue(1)
                .AddFalse(2, 3)
                .AssertVisibility();
        }

        [Test]
        public void AboveAverage_takes_values_over_avg_value()
        {
            new AutoFilterTester(f => f.AboveAverage())
                .AddTrue(3)
                .AddFalse(2, 1)
                .AssertVisibility();
        }

        [Test]
        public void Average_ignores_non_unified_numbers()
        {
            new AutoFilterTester(f => f.BelowAverage())
                .AddTrue(new DateTime(1900, 1, 1)) // Serial date time 1
                .AddFalse(1.1)
                .AddFalse(1.2)
                .AddFalse(XLError.NoValueAvailable, true, false, "-100", "Text", Blank.Value)
                .AssertVisibility();
        }

        [Test]
        public void All_rows_are_hidden_when_column_has_no_number()
        {
            new AutoFilterTester(f => f.AboveAverage())
                .AddFalse(Blank.Value, true, false, "-100", "Text", XLError.NoValueAvailable)
                .AssertVisibility();
        }
    }
}
