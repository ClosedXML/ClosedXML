using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML_Tests.Excel.Formula
{
    public class XLFormulaDefinitionParserTests
    {
        [TestCase("RC", typeof(XLCellReference))]
        [TestCase("RC1", typeof(XLCellReference))]
        [TestCase("R2C3", typeof(XLCellReference))]
        [TestCase("R[-2]C[-3]", typeof(XLCellReference))]
        [TestCase("RC:R[-2]C[-3]", typeof(XLRangeReference))]

        [TestCase("R", typeof(XLRowReference))]
        [TestCase("R1", typeof(XLRowReference))]
        [TestCase("R[-2]", typeof(XLRowReference))]

        [TestCase("R:R", typeof(XLRowRangeReference))]
        [TestCase("R[-1]:R", typeof(XLRowRangeReference))]
        [TestCase("R2:R3", typeof(XLRowRangeReference))]

        [TestCase("C", typeof(XLColumnReference))]
        [TestCase("C1", typeof(XLColumnReference))]
        [TestCase("C[-2]", typeof(XLColumnReference))]

        [TestCase("C:C", typeof(XLColumnRangeReference))]
        [TestCase("C[-1]:C", typeof(XLColumnRangeReference))]
        [TestCase("C2:C3", typeof(XLColumnRangeReference))]
        public void ParseCorrectTypesR1C1(string referenceString, Type expectedReferenceType)
        {
            var parser = new XLFormulaDefinitionR1C1Parser();

            var res = parser.Parse(referenceString);

            Assert.AreEqual(1, res.Item2.Length);
            Assert.IsAssignableFrom(expectedReferenceType, res.Item2[0]);
            Assert.AreEqual(referenceString, res.Item2[0].ToStringR1C1());
        }
    }
}
