using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class ArithmeticOperatorsTests
    {
        #region Concat text operator

        [TestCase("\"A\" & \"B\"", "AB")]
        [TestCase("\"\" & \"B\"", "B")]
        [TestCase("\"A\" & \"\"", "A")]
        [TestCase("\"\" & \"\"", "")]
        public void Concat_ConcatenateText(string formula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("A1 & \"\"", "")]
        [TestCase("\"\" & A1", "")]
        [TestCase("A1 & A1", "")]
        public void Concat_ConcatenateBlank(string formula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, Evaluate(formula));
        }

        [TestCase("TRUE & \" to text\"", "TRUE to text")]
        [TestCase("FALSE & \" to text\"", "FALSE to text")]
        [TestCase("true & \" to text\"", "TRUE to text")]
        [TestCase("false & \" to text\"", "FALSE to text")]
        [TestCase("TRUE & FALSE", "TRUEFALSE")]
        public void Concat_ConvertsLogicalToString(string formula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        [SetCulture("cs-CZ")]
        [TestCase("1 & \" to text\"", "1 to text")]
        [TestCase("1 & 0", "10")]
        [TestCase("1.5 & 0.78", "1,50,78")]
        public void Concat_ConvertsNumberToStringUsingCulture(string formula, object expectedResult)
        {
            var wb = new XLWorkbook();
            Assert.AreEqual(expectedResult, wb.Evaluate(formula));
        }

        [TestCase("#DIV/0! & 1", XLError.DivisionByZero)]
        [TestCase("#DIV/0! & \"1\"", XLError.DivisionByZero)]
        [TestCase("#REF! & #DIV/0!", XLError.CellReference)]
        [TestCase("1 & #NAME?", XLError.NameNotRecognized)]
        public void Concat_WithErrorAsOperandReturnsTheError(string formula, XLError expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr(formula));
        }

        [Ignore("Arrays are not implemented")]
        [TestCase("{1,2} & \"A\"", "1A")]
        [TestCase("{\"A\",2} & \"B\"", "AB")]
        [TestCase("{TRUE,2} & \"B\"", "TRUEB")]
        [TestCase("{#REF!,5} & 1", XLError.CellReference)]
        public void Concat_UsesFirstElementOfArray(string formula, object expected)
        {
            Assert.AreEqual(expected, XLWorkbook.EvaluateExpr(formula));
        }

        #endregion

        #region Unary plus

        [TestCase("+1", 1)]
        [TestCase("+\"1\"", "1")]
        [TestCase("+TRUE", true)]
        [TestCase("+FALSE", false)]
        [TestCase("+#DIV/0!", XLError.DivisionByZero)]
        [TestCase("+A1", 0)]
        public void UnaryPlus_IsNonOpThatKeepsValueAndType(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Unary minus

        [TestCase("-1", -1)]
        [TestCase("-125.45", -125.45)]
        [TestCase("-\"1\"", -1)]
        [TestCase("-TRUE", -1)]
        [TestCase("-FALSE", 0)]
        [TestCase("-#DIV/0!", XLError.DivisionByZero)]
        [TestCase("-A1", 0.0)]
        public void UnaryMinus_ConvertsArgumentBeforeNegating(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Unary minus

        [TestCase("1%", 0.01)]
        [TestCase("100%", 1.0)]
        [TestCase("25.7%", 0.257)]
        [TestCase("125.45%", 1.2545)]
        [TestCase("\"1\"%", 0.01)]
        [TestCase("TRUE%", 0.01)]
        [TestCase("FALSE%", 0)]
        [TestCase("#NAME?%", XLError.NameNotRecognized)]
        [TestCase("(1/0)%", XLError.DivisionByZero)]
        [TestCase("A1%", 0.0)]
        public void UnaryPercent_ConvertsArgumentBeforePercentOperator(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Exponentiation

        [TestCase("1^1", 1.0)]
        [TestCase("0^0", XLError.NumberInvalid)]
        [TestCase("10^0", 1.0)]
        [TestCase("4^0.5", 2.0)]
        [TestCase("2^0.5", 1.4142135623730951)]
        [TestCase("2^-2", 0.25)]
        [TestCase("\"5\"^\"3\"", 125)]
        [TestCase("5^TRUE", 5)]
        [TestCase("5^FALSE", 1)]
        [TestCase("#VALUE!^1", XLError.IncompatibleValue)]
        [TestCase("1^#REF!", XLError.CellReference)]
        [TestCase("#DIV/0!^#REF!", XLError.DivisionByZero)]
        [TestCase("5^A1", 1.0)]
        [TestCase("A1^4", 0.0)]
        public void Exponentiation_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.That(Evaluate(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
        }

        #endregion

        #region Multiplication

        [TestCase("1+1", 2.0)]
        [TestCase("0*0", 0.0)]
        [TestCase("10*0", 0.0)]
        [TestCase("2*1.5", 3.0)]
        [TestCase("2.5*2.5", 6.25)]
        [TestCase("2*-2", -4)]
        [TestCase("\"5\" * \"3\"", 15)]
        [TestCase("5*TRUE", 5)]
        [TestCase("5*FALSE", 0)]
        [TestCase("#VALUE!*1", XLError.IncompatibleValue)]
        [TestCase("1*#REF!", XLError.CellReference)]
        [TestCase("#DIV/0!*#REF!", XLError.DivisionByZero)]
        [TestCase("10*A1", 0.0)]
        [TestCase("A1*10", 0.0)]
        public void Multiplication_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.That(Evaluate(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
        }

        #endregion

        #region Division

        [TestCase("1/1", 1.0)]
        [TestCase("5/2", 2.5)]
        [TestCase("14.5/2.5", 5.8)]
        [TestCase("10/0", XLError.DivisionByZero)]
        [TestCase("0/0", XLError.DivisionByZero)]
        [TestCase("2.5/-0.5", -5)]
        [TestCase("\"10\" / \"4\"", 2.5)]
        [TestCase("5/TRUE", 5)]
        [TestCase("5/FALSE", XLError.DivisionByZero)]
        [TestCase("#VALUE!/1", XLError.IncompatibleValue)]
        [TestCase("1/#REF!", XLError.CellReference)]
        [TestCase("#DIV/0!/#REF!", XLError.DivisionByZero)]
        [TestCase("A1/5", 0.0)]
        [TestCase("5/A1", XLError.DivisionByZero)]
        public void Division_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Addition

        [TestCase("1+1", 2.0)]
        [TestCase("5+2.5", 7.5)]
        [TestCase("10+0", 10.0)]
        [TestCase("\"10\" + \"4\"", 14.0)]
        [TestCase("5+TRUE", 6.0)]
        [TestCase("5+FALSE", 5.0)]
        [TestCase("#VALUE! + 1", XLError.IncompatibleValue)]
        [TestCase("1 + #REF!", XLError.CellReference)]
        [TestCase("#DIV/0! + #REF!", XLError.DivisionByZero)]
        [TestCase("A1 + 7", 7)]
        public void Addition_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Subtraction

        [TestCase("1-1", 0.0)]
        [TestCase("2.5-7.8", -5.3)]
        [TestCase("10-0", 10.0)]
        [TestCase("\"10\" - \"4\"", 6.0)]
        [TestCase("5-TRUE", 4.0)]
        [TestCase("5-FALSE", 5.0)]
        [TestCase("#VALUE! - 1", XLError.IncompatibleValue)]
        [TestCase("1 - #REF!", XLError.CellReference)]
        [TestCase("#DIV/0! - #REF!", XLError.DivisionByZero)]
        [TestCase("A1 - 5", -5)]
        public void Subtraction_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        private static object Evaluate(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            return ws.Evaluate(formula);
        }
    }
}
