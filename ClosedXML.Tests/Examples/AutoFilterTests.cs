using ClosedXML.Examples;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class AutoFilterTests
    {
        [Test]
        public void CustomAutoFilter()
        {
            TestHelper.RunTestExample<CustomAutoFilter>(@"AutoFilter\CustomAutoFilter.xlsx");
        }

        [Test]
        public void DynamicAutoFilter()
        {
            TestHelper.RunTestExample<DynamicAutoFilter>(@"AutoFilter\DynamicAutoFilter.xlsx");
        }

        [Test]
        public void RegularAutoFilter()
        {
            TestHelper.RunTestExample<RegularAutoFilter>(@"AutoFilter\RegularAutoFilter.xlsx");
        }

        [Test]
        public void TopBottomAutoFilter()
        {
            TestHelper.RunTestExample<TopBottomAutoFilter>(@"AutoFilter\TopBottomAutoFilter.xlsx");
        }

        [Test]
        public void DateTimeGroupAutoFilter()
        {
            TestHelper.RunTestExample<DateTimeGroupAutoFilter>(@"AutoFilter\DateTimeGroupAutoFilter.xlsx");
        }
    }
}
