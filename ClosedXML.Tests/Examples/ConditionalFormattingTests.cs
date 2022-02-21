using ClosedXML.Examples;
using NUnit.Framework;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class ConditionalFormattingTests
    {
        [Test]
        public void CFColorScaleLowHigh()
        {
            TestHelper.RunTestExample<CFColorScaleLowHigh>(@"ConditionalFormatting\CFColorScaleLowHigh.xlsx");
        }

        [Test]
        public void CFColorScaleLowMidHigh()
        {
            TestHelper.RunTestExample<CFColorScaleLowMidHigh>(@"ConditionalFormatting\CFColorScaleLowMidHigh.xlsx");
        }

        [Test]
        public void CFColorScaleMinimumMaximum()
        {
            TestHelper.RunTestExample<CFColorScaleMinimumMaximum>(@"ConditionalFormatting\CFColorScaleMinimumMaximum.xlsx");
        }

        [Test]
        public void CFContains()
        {
            TestHelper.RunTestExample<CFContains>(@"ConditionalFormatting\CFContains.xlsx");
        }

        [Test]
        public void CFDataBar()
        {
            TestHelper.RunTestExample<CFDataBar>(@"ConditionalFormatting\CFDataBar.xlsx");
        }

        [Test]
        public void CFDataBarNegative()
        {
            TestHelper.RunTestExample<CFDataBarNegative>(@"ConditionalFormatting\CFDataBarNegative.xlsx");
        }

        [Test]
        public void CFEndsWith()
        {
            TestHelper.RunTestExample<CFEndsWith>(@"ConditionalFormatting\CFEndsWith.xlsx");
        }

        [Test]
        public void CFEqualsNumber()
        {
            TestHelper.RunTestExample<CFEqualsNumber>(@"ConditionalFormatting\CFEqualsNumber.xlsx");
        }

        [Test]
        public void CFEqualsString()
        {
            TestHelper.RunTestExample<CFEqualsString>(@"ConditionalFormatting\CFEqualsString.xlsx");
        }

        [Test]
        public void CFIconSet()
        {
            TestHelper.RunTestExample<CFIconSet>(@"ConditionalFormatting\CFIconSet.xlsx");
        }

        [Test]
        public void CFIsBlank()
        {
            TestHelper.RunTestExample<CFIsBlank>(@"ConditionalFormatting\CFIsBlank.xlsx");
        }

        [Test]
        public void CFIsError()
        {
            TestHelper.RunTestExample<CFIsError>(@"ConditionalFormatting\CFIsError.xlsx");
        }

        [Test]
        public void CFNotBlank()
        {
            TestHelper.RunTestExample<CFNotBlank>(@"ConditionalFormatting\CFNotBlank.xlsx");
        }

        [Test]
        public void CFNotContains()
        {
            TestHelper.RunTestExample<CFNotContains>(@"ConditionalFormatting\CFNotContains.xlsx");
        }

        [Test]
        public void CFNotEqualsNumber()
        {
            TestHelper.RunTestExample<CFNotEqualsNumber>(@"ConditionalFormatting\CFNotEqualsNumber.xlsx");
        }

        [Test]
        public void CFNotEqualsString()
        {
            TestHelper.RunTestExample<CFNotEqualsString>(@"ConditionalFormatting\CFNotEqualsString.xlsx");
        }

        [Test]
        public void CFNotError()
        {
            TestHelper.RunTestExample<CFNotError>(@"ConditionalFormatting\CFNotError.xlsx");
        }

        [Test]
        public void CFStartsWith()
        {
            TestHelper.RunTestExample<CFStartsWith>(@"ConditionalFormatting\CFStartsWith.xlsx");
        }

        [Test]
        public void CFMultipleConditions()
        {
            TestHelper.RunTestExample<CFMultipleConditions>(@"ConditionalFormatting\CFMultipleConditions.xlsx");
        }

        [Test]
        public void CFStopIfTrue()
        {
            TestHelper.RunTestExample<CFStopIfTrue>(@"ConditionalFormatting\CFStopIfTrue.xlsx");
        }

        [Test]
        public void CFTop()
        {
            TestHelper.RunTestExample<CFTop>(@"ConditionalFormatting\CFTop.xlsx");
        }

        [Test]
        public void CFBottom()
        {
            TestHelper.RunTestExample<CFBottom>(@"ConditionalFormatting\CFBottom.xlsx");
        }

        [Test]
        public void CFDatesOccurring()
        {
            TestHelper.RunTestExample<CFDatesOccurring>(@"ConditionalFormatting\CFDatesOccurring.xlsx");
        }

        [Test]
        public void CFDataBars()
        {
            TestHelper.RunTestExample<CFDataBars>(@"ConditionalFormatting\CFDataBars.xlsx");
        }
    }
}
