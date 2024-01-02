using System;
using NUnit.Framework;
using System.Linq;
using ClosedXML.Excel;

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

        [Test]
        public void Top_items_filter_excludes_non_unified_numbers()
        {
            // Sort and then use cutoff value, it's 4 here and then take all values >= cutoff.
            new AutoFilterTester(f => f.Top(1))
                .AddTrue(new DateTime(1900, 2, 10))
                .AddFalse(11, 10)
                .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
                .AssertVisibility();
        }

        [Test]
        public void Bottom_items_filter_excludes_non_unified_numbers()
        {
            new AutoFilterTester(f => f.Bottom(1))
                .AddTrue(new DateTime(1900, 1, 1))
                .AddFalse(2, 3)
                .AddFalse("-1000", "Text", Blank.Value, true, false, XLError.IncompatibleValue)
                .AssertVisibility();
        }

        [Test]
        public void Top_items_filter_determines_top_items_by_determining_cut_off_value()
        {
            // Sort and then use cutoff value, it's 4 here and then take all values <= cutoff.
            new AutoFilterTester(f => f.Top(2))
                .AddTrue(5, 4, 4, 4)
                .AddFalse(3, 2, 1)
                .AssertVisibility();

            // Cutoff is 5 here.
            new AutoFilterTester(f => f.Top(2))
                .AddTrue(5, 5)
                .AddFalse(4, 4, 4, 3, 2, 1)
                .AssertVisibility();
        }

        [Test]
        public void Bottom_items_filter_determines_top_items_by_determining_cut_off_value()
        {
            // Cutoff is 2
            new AutoFilterTester(f => f.Bottom(2))
                .AddTrue(1, 2, 2, 2)
                .AddFalse(3, 4, 5)
                .AssertVisibility();

            // Cutoff is 5
            new AutoFilterTester(f => f.Bottom(2))
                .AddTrue(1, 1)
                .AddFalse(2, 2, 2, 3, 4, 5)
                .AssertVisibility();
        }

        [Test]
        public void Top_percents_uses_inclusive_percent_value()
        {
            // Autofilter doesn't include value 750, which is at 75%, i.e. right at the border.
            new AutoFilterTester(f => f.Top(25, XLTopBottomType.Percent))
                .AddFalse(Enumerable.Range(1, 750).Select<int, XLCellValue>(x => x).ToArray())
                .AddTrue(Enumerable.Range(751, 250).Select<int, XLCellValue>(x => x).ToArray())
                .AssertVisibility();
        }

        [Test]
        public void Bottom_percents_uses_inclusive_percent_value()
        {
            new AutoFilterTester(f => f.Bottom(25, XLTopBottomType.Percent))
                .AddTrue(Enumerable.Range(1, 250).Select<int, XLCellValue>(x => x).ToArray())
                .AddFalse(Enumerable.Range(251, 750).Select<int, XLCellValue>(x => x).ToArray())
                .AssertVisibility();
        }

        [Test]
        public void Top_percents_always_has_at_least_one_item()
        {
            // Top 1% takes one item that is 33% of all items.
            new AutoFilterTester(f => f.Top(1, XLTopBottomType.Percent))
                .AddTrue(3)
                .AddFalse(2, 1)
                .AssertVisibility();
        }

        [Test]
        public void Bottom_percents_always_has_at_least_one_item()
        {
            new AutoFilterTester(f => f.Bottom(1, XLTopBottomType.Percent))
                .AddTrue(1)
                .AddFalse(2, 3)
                .AssertVisibility();
        }

        [TestCase(0, true)]
        [TestCase(501, true)]
        [TestCase(0, false)]
        [TestCase(501, false)]
        public void Top_and_bottom_filter_value_must_be_between_1_and_500(int value, bool top)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Data";
            ws.Cell("A2").Value = value;
            var autoFilter = ws.Range("A1:A2").SetAutoFilter();
            var filterColumn = autoFilter.Column(1);

            var ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
            {
                if (top)
                    filterColumn.Top(value);
                else
                    filterColumn.Bottom(value);
            })!;
            StringAssert.Contains("Value must be between 1 and 500.", ex.Message);
        }
    }
}
