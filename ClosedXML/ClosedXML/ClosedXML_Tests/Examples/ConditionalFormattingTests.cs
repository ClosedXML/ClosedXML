using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class ConditionalFormattingTests
    {
    
        [TestMethod]
        public void CFColorScaleLowMidHigh()
        {
            TestHelper.RunTestExample<CFColorScaleLowMidHigh>(@"ConditionalFormatting\CFColorScaleLowMidHigh.xlsx");
        }

        [TestMethod]
        public void CFColorScaleLowHigh()
        {
            TestHelper.RunTestExample<CFColorScaleLowHigh>(@"ConditionalFormatting\CFColorScaleLowHigh.xlsx");
        }

        [TestMethod]
        public void CFStartsWith()
        {
            TestHelper.RunTestExample<CFStartsWith>(@"ConditionalFormatting\CFStartsWith.xlsx");
        }

        [TestMethod]
        public void CFEndsWith()
        {
            TestHelper.RunTestExample<CFEndsWith>(@"ConditionalFormatting\CFEndsWith.xlsx");
        }

        [TestMethod]
        public void CFIsBlank()
        {
            TestHelper.RunTestExample<CFIsBlank>(@"ConditionalFormatting\CFIsBlank.xlsx");
        }

        [TestMethod]
        public void CFNotBlank()
        {
            TestHelper.RunTestExample<CFNotBlank>(@"ConditionalFormatting\CFNotBlank.xlsx");
        }

        [TestMethod]
        public void CFIsError()
        {
            TestHelper.RunTestExample<CFIsError>(@"ConditionalFormatting\CFIsError.xlsx");
        }

        [TestMethod]
        public void CFNotError()
        {
            TestHelper.RunTestExample<CFNotError>(@"ConditionalFormatting\CFNotError.xlsx");
        }

        [TestMethod]
        public void CFContains()
        {
            TestHelper.RunTestExample<CFContains>(@"ConditionalFormatting\CFContains.xlsx");
        }

        [TestMethod]
        public void CFNotContains()
        {
            TestHelper.RunTestExample<CFNotContains>(@"ConditionalFormatting\CFNotContains.xlsx");
        }

        [TestMethod]
        public void CFEqualsString()
        {
            TestHelper.RunTestExample<CFEqualsString>(@"ConditionalFormatting\CFEqualsString.xlsx");
        }

        [TestMethod]
        public void CFEqualsNumber()
        {
            TestHelper.RunTestExample<CFEqualsNumber>(@"ConditionalFormatting\CFEqualsNumber.xlsx");
        }

        [TestMethod]
        public void CFNotEqualsString()
        {
            TestHelper.RunTestExample<CFNotEqualsString>(@"ConditionalFormatting\CFNotEqualsString.xlsx");
        }

        [TestMethod]
        public void CFNotEqualsNumber()
        {
            TestHelper.RunTestExample<CFNotEqualsNumber>(@"ConditionalFormatting\CFNotEqualsNumber.xlsx");
        }

        [TestMethod]
        public void CFDataBar()
        {
            TestHelper.RunTestExample<CFDataBar>(@"ConditionalFormatting\CFDataBar.xlsx");
        }

        [TestMethod]
        public void CFIconSet()
        {
            TestHelper.RunTestExample<CFIconSet>(@"ConditionalFormatting\CFIconSet.xlsx");
        }
        //
        //[TestMethod]
        //public void XXX()
        //{
        //    TestHelper.RunTestExample<XXX>(@"ConditionalFormatting\XXX.xlsx");
        //}
        //
        //[TestMethod]
        //public void XXX()
        //{
        //    TestHelper.RunTestExample<XXX>(@"ConditionalFormatting\XXX.xlsx");
        //}
        //
        //[TestMethod]
        //public void XXX()
        //{
        //    TestHelper.RunTestExample<XXX>(@"ConditionalFormatting\XXX.xlsx");
        //}
        //
        //[TestMethod]
        //public void XXX()
        //{
        //    TestHelper.RunTestExample<XXX>(@"ConditionalFormatting\XXX.xlsx");
        //}
    }
}