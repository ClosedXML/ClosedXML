using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class AutoFilterTests
    {
    
    [TestMethod]
        public void RegularAutoFilter()
        {
            TestHelper.RunTestExample<RegularAutoFilter>(@"AutoFilter\RegularAutoFilter.xlsx");
        }
    [TestMethod]
        public void CustomAutoFilter()
        {
            TestHelper.RunTestExample<CustomAutoFilter>(@"AutoFilter\CustomAutoFilter.xlsx");
        }

    [TestMethod]
    public void TopBottomAutoFilter()
        {
            TestHelper.RunTestExample<TopBottomAutoFilter>(@"AutoFilter\TopBottomAutoFilter.xlsx");
        }

    [TestMethod]
    public void DynamicAutoFilter()
        {
            TestHelper.RunTestExample<DynamicAutoFilter>(@"AutoFilter\DynamicAutoFilter.xlsx");
        }
    }
}