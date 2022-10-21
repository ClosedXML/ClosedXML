using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CompareOperatorsTests
    {
        [TestCase("1=1", true)]
        [TestCase("1=0", false)]
        [TestCase("0.0=0", true)]
        [TestCase("TRUE=TRUE", true)]
        [TestCase("FALSE=FALSE", true)]
        [TestCase("TRUE=FALSE", false)]
        [TestCase("\"text\"=\"text\"", true)]
        [TestCase("\"tExT\"=\"TeXt\"", true)]
        [TestCase("\"text\"=\"text\"", true)]
        [TestCase("\"\"=\"\"", true)]
        [TestCase("#VALUE!=#VALUE!", XLError.IncompatibleValue)]
        [TestCase("A1=B1", true)] // blanks are equal
        public void EqualTo_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("1<>1", false)]
        [TestCase("1<>0", true)]
        [TestCase("0.0<>0", false)]
        [TestCase("TRUE<>TRUE", false)]
        [TestCase("FALSE<>FALSE", false)]
        [TestCase("TRUE<>FALSE", true)]
        [TestCase("\"texty\"<>\"text\"", true)]
        [TestCase("\"tExT\"<>\"TeXt\"", false)]
        [TestCase("\"text\"<>\"text\"", false)]
        [TestCase("\"\"<>\"\"", false)]
        [TestCase("#VALUE!<>#VALUE!", XLError.IncompatibleValue)]
        [TestCase("A1<>B1", false)] // blanks are equal
        public void NotEqualTo_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("1>1", false)]
        [TestCase("1>0", true)]
        [TestCase("0.0>0", false)]
        [TestCase("TRUE>TRUE", false)]
        [TestCase("FALSE>FALSE", false)]
        [TestCase("TRUE>FALSE", true)]
        [TestCase("\"text\">\"text\"", false)]
        [TestCase("\"texu\">\"text\"", true)]
        [TestCase("#VALUE!>#REF!", XLError.IncompatibleValue)]
        [TestCase("A1>A2", false)]
        public void GreaterThen_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("1>=1", true)]
        [TestCase("1>=0", true)]
        [TestCase("0.0>=0", true)]
        [TestCase("TRUE>=TRUE", true)]
        [TestCase("FALSE>=FALSE", true)]
        [TestCase("TRUE>=FALSE", true)]
        [TestCase("\"text\">=\"text\"", true)]
        [TestCase("\"texu\">=\"text\"", true)]
        [TestCase("#VALUE!>=#REF!", XLError.IncompatibleValue)]
        [TestCase("A1>=A2", true)]
        public void GreaterThenOrEqual_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("-5<5", true)]
        [TestCase("1<1", false)]
        [TestCase("1<0", false)]
        [TestCase("0.0<0", false)]
        [TestCase("TRUE<TRUE", false)]
        [TestCase("FALSE<FALSE", false)]
        [TestCase("TRUE<FALSE", false)]
        [TestCase("FALSE<TRUE", true)]
        [TestCase("\"text\"<\"text\"", false)]
        [TestCase("\"text\"<\"texu\"", true)]
        [TestCase("#VALUE!<#REF!", XLError.IncompatibleValue)]
        [TestCase("A1<A2", false)]
        public void LessThen_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("-5<=5", true)]
        [TestCase("1<=1", true)]
        [TestCase("1<=0", false)]
        [TestCase("0.0<=0", true)]
        [TestCase("TRUE<=TRUE", true)]
        [TestCase("FALSE<=FALSE", true)]
        [TestCase("TRUE<=FALSE", false)]
        [TestCase("FALSE<=TRUE", true)]
        [TestCase("\"text\"<=\"text\"", true)]
        [TestCase("\"text\"<=\"texu\"", true)]
        [TestCase("#VALUE!<=#REF!", XLError.IncompatibleValue)]
        [TestCase("A1<=A2", true)]
        public void LessThenOrEqual_WithSameType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        [TestCase("TRUE>-1", true)]
        [TestCase("TRUE>1", true)]
        [TestCase("TRUE>100", true)]
        [TestCase("FALSE>-1", true)]
        [TestCase("FALSE>1", true)]
        [TestCase("FALSE>100", true)]
        [TestCase("TRUE>\"100\"", true)]
        [TestCase("FALSE>\"100\"", true)]
        [TestCase("FALSE>\"\"", true)]
        [TestCase("\"\">FALSE", false)]
        [TestCase("10>FALSE", false)]
        [TestCase("10>TRUE", false)]
        [TestCase("-1<TRUE", true)]
        [TestCase("1<TRUE", true)]
        [TestCase("100<TRUE", true)]
        [TestCase("-1<FALSE", true)]
        [TestCase("1<FALSE", true)]
        [TestCase("100<FALSE", true)]
        [TestCase("\"100\"<TRUE", true)]
        [TestCase("\"100\"<FALSE", true)]
        [TestCase("\"\"<FALSE", true)]
        [TestCase("FALSE<\"\"", false)]
        [TestCase("FALSE<10", false)]
        [TestCase("TRUE<10", false)]
        public void Comparison_LogicalIsAlwaysGreaterThanAnyTextOrNumber(string formula, bool expectedResult)
        {
            Assert.AreEqual(expectedResult, Evaluate(formula));
        }

        [TestCase("\"\">10", true)]
        [TestCase("\"1\">10", true)]
        [TestCase("10<\"\"", true)]
        [TestCase("10<\"1\"", true)]
        public void Comparison_TextIsAlwaysGreaterThanAnyNumber(string formula, bool expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("FALSE=A1")]
        [TestCase("A1=FALSE")]
        [TestCase("A1=0")]
        [TestCase("0=A1")]
        [TestCase("\"\"=A1")]
        [TestCase("A1=\"\"")]
        public void Comparison_BlankIsEqualToFalseOrZeroOrEmptyString(string formula)
        {
            Assert.That(Evaluate(formula), Is.True);
        }

        private static object Evaluate(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            return ws.Evaluate(formula);
        }
    }
}
