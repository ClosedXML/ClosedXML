using ClosedXML.Excel;
using NUnit.Framework;
using System;

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
        [TestCase("TRUE & FALSE", @"TRUEFALSE")]
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

        #endregion

        #region Unary plus

        [TestCase("+1", 1)]
        [TestCase("+\"1\"", "1")]
        [TestCase("+TRUE", true)]
        [TestCase("+FALSE", false)]
        [TestCase("+#DIV/0!", XLError.DivisionByZero)]
        [TestCase("ISBLANK(+A1)", true)]
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

        #region Array Operations

        [Test]
        public void ArraysOperation_BinaryOperationBetweenAreaReferenceAndSingleCellReferenceShouldWork()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Test1");
            ws.Cell("A1").Value = new DateTime(2021, 1, 15);
            ws.Cell("A2").Value = new DateTime(2021, 1, 10);
            ws.Cell("B1").Value = new DateTime(2021, 1, 5);
            Assert.AreEqual(5, ws.Evaluate("MIN(A1:A2-B1)"));
        }

        [Test]
        public void ArraysOperation_MultiAreaReferencesArgumentResultsInScalarError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cells("A1:A2").Value = 1;
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("(A1:A1,A1:A2)+1"));
            Assert.AreEqual(16, ws.Evaluate("TYPE((A1:A1,A1:A2)+1)")); // The result is a scalar error, not an array of errors
        }

        [Test]
        public void ArrayOperation_SameSizeArrayPerformsOperationIndividually()
        {
            Assert.AreEqual(6 * 7, XLWorkbook.EvaluateExpr("SUM({1,2,3;4,5,6} + {6,5,4;3,2,1})"));
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("COLUMNS({1,2} + \"A\")"));
        }

        [Test]
        public void ArrayOperation_ArrayPlusScalarUpscalesScalarToSizeOfArray()
        {
            Assert.AreEqual(18, XLWorkbook.EvaluateExpr("SUM({1,1,1;1,1,1} * 3)"));
            Assert.AreEqual(15, XLWorkbook.EvaluateExpr("SUM(6 / {2,2,2;3,3,3})"));
        }

        [Test]
        public void ArrayOperation_RowOnlyArrayIsRepeatedToHaveSameNumberOfRowsAsOtherArray()
        {
            // {3,2} is scaled to {3,2;3,2} of second array
            Assert.AreEqual(14, XLWorkbook.EvaluateExpr("SUM({3,2}+{1,1;1,1})"));
            Assert.AreEqual(14, XLWorkbook.EvaluateExpr("SUM({1,1;1,1}+{3,2})"));
        }

        [Test]
        public void ArrayOperation_ColumnOnlyArrayIsRepeatedToHaveSameNumberOfColumnsAsOtherArray()
        {
            // {3;2} is scaled to {3,3;2,2} of second array
            Assert.AreEqual(16, XLWorkbook.EvaluateExpr("SUM({3;2}*{1,1;2,3})"));
            Assert.AreEqual(16, XLWorkbook.EvaluateExpr("SUM({1,1;2,3}*{3;2})"));
        }

        [Test]
        public void ArrayOperation_1x1ArrayIsScaledToOtherArray()
        {
            Assert.AreEqual(20, XLWorkbook.EvaluateExpr("SUM({2}*{1,2;3,4})"));
            Assert.AreEqual(20, XLWorkbook.EvaluateExpr("SUM({1,2;3,4}*{2})"));
        }

        [Test]
        public void ArrayOperation_DifferentSizedArraysAreUpscaledToContainingSize()
        {
            // The extra value are #N/A + value, i.e. #N/A, thus the whole sum is #N/A
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("SUM({1,2;3,4;5,6}+{1,2,3;4,5,6})"));
            Assert.AreEqual(3, XLWorkbook.EvaluateExpr("ROWS({1,2;3,4;5,6}+{1,2,3;4,5,6})"));
            Assert.AreEqual(3, XLWorkbook.EvaluateExpr("COLUMNS({1,2;3,4;5,6}+{1,2,3;4,5,6})"));
        }

        #endregion

        private static XLCellValue Evaluate(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            return ws.Evaluate(formula);
        }
    }
}
