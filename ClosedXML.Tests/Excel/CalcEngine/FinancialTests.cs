using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FinancialTests
    {
        private readonly double tolerance = 1e-10;

        [TestCase("PMT(0.08/12,10,10000)", -1037.03208935915)]
        [TestCase("PMT(0.08/12,10,10000,0,1)", -1030.16432717797)]
        public void Pmt_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [Test]
        public void Pmt_PaymentsMustPayForPrincipalAndFutureValue()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("PMT(0,2,5000,10000)");
            Assert.AreEqual(-7500, actual);
        }

        [TestCase("PMT(0,1,1000)", -1000)] // Zero interest rate
        [TestCase("PMT(0,5,10000,5000)", -3000)] // Zero interest rate for 5 years, (10k principal, pay all and have 5k in bank at the end = payment is 3k/year)
        [TestCase("PMT(-0.4,2,1000)", -225)] // Negative interest rate
        [TestCase("PMT(0.01,0.5,1000)", -2014.98756211209)] // Non-integer period
        [TestCase("PMT(0.1,-2,1000)", 476.19047619048)] // Negative periods
        [TestCase("PMT(0.1,2,0,4)", -1.90476190476)] // No principal, but future value
        [TestCase("PMT(0,2,-1000)", 500)] // Negative principal - money is paid to us
        [TestCase("PMT(0.000001,1000,1000)", -1.00050058333321)] // Small number and high number of periods, check for stability
        public void Pmt_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, tolerance);
        }

        [Test]
        public void Pmt_TypeConvertsNumberToOneOrZero()
        {
            // Spec says "if type is any number other than 0 or 1, #NUM! is returned.", but Excel accepts any number as type
            var formulaFormat = "PMT(0.1,2,1000,500,{0})";
            var zeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0"));
            var oneType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "1"));
            var nonZeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0.000001"));

            Assert.AreNotEqual(zeroType, oneType);
            Assert.AreEqual(oneType, nonZeroType);
        }

        [Test]
        public void Pmt_DefaultFutureValueIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("PMT(0.1,2,1000)"), XLWorkbook.EvaluateExpr("PMT(0.1,2,1000,0)"));
        }

        [Test]
        public void Pmt_DefaultTypeIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("PMT(0.1,5,1000)"), XLWorkbook.EvaluateExpr("PMT(0.1,5,1000,0,0)"));
        }

        [Test]
        public void Pmt_ZeroPeriodsReturnsNumError()
        {
            Assert.Throws<NumberException>(() => XLWorkbook.EvaluateExpr("PMT(0.1,0,1000)"));
        }
    }
}
