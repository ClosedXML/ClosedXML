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

            var res = parser.Parse(referenceString, null);

            Assert.AreEqual(1, res.Item2.Length);
            Assert.IsAssignableFrom(expectedReferenceType, res.Item2[0]);
            Assert.AreEqual(referenceString, res.Item2[0].ToStringR1C1());
        }

        [TestCase("D8", typeof(XLCellReference))]
        [TestCase("$A8", typeof(XLCellReference))]
        [TestCase("$C$2", typeof(XLCellReference))]
        [TestCase("$C7", typeof(XLCellReference))]
        [TestCase("B3", typeof(XLCellReference))]
        [TestCase("$B$3:F6", typeof(XLRangeReference))]
        [TestCase("B2:C$3", typeof(XLRangeReference))]

        [TestCase("2:2", typeof(XLRowRangeReference))]
        [TestCase("$4:5", typeof(XLRowRangeReference))]
        [TestCase("5:$6", typeof(XLRowRangeReference))]
        [TestCase("$1:$4", typeof(XLRowRangeReference))]

        [TestCase("B:D", typeof(XLColumnRangeReference))]
        [TestCase("$B:D", typeof(XLColumnRangeReference))]
        [TestCase("B:$D", typeof(XLColumnRangeReference))]
        [TestCase("$B:$D", typeof(XLColumnRangeReference))]
        public void ParseCorrectTypesA1(string referenceString, Type expectedReferenceType)
        {
            var parser = new XLFormulaDefinitionA1Parser();
            var baseAddress = new XLAddress(8, 4, false, false);
            var res = parser.Parse(referenceString, baseAddress);

            Assert.AreEqual(1, res.Item2.Length);
            Assert.IsAssignableFrom(expectedReferenceType, res.Item2[0]);
            Assert.AreEqual(referenceString, res.Item2[0].ToStringA1(baseAddress));
        }

        [TestCase("=\"\"", 0)]
        [TestCase("=COLUMN()", 0)]
        [TestCase("=COLUMN(RC)", 1)]
        [TestCase("RC+12", 1)]
        [TestCase("12+RC-34", 1)]
        [TestCase("12+RC", 1)]
        [TestCase("12+RC+SUM(R[-1]:R[1])", 2)]
        [TestCase("RC1*RC2*RC3*RC4*R5C", 5)]
        [TestCase("='S10 Data'!R1C1", 1)]
        public void ExtractReferencesFromFormula(string formula, int expectedNumberOfReferences)
        {
            var parser = new XLFormulaDefinitionR1C1Parser();

            var res = parser.Parse(formula, null);

            Assert.AreEqual(expectedNumberOfReferences, res.Item2.Length);
            Assert.AreEqual(expectedNumberOfReferences + 1, res.Item1.Length);

            var restoredFormula = "";
            for (int i = 0; i < expectedNumberOfReferences; i++)
            {
                restoredFormula += res.Item1[i];
                restoredFormula += res.Item2[i];
            }

            restoredFormula += res.Item1[expectedNumberOfReferences];

            Assert.AreEqual(formula, restoredFormula);
        }
    }
}
