using System;
using System.Globalization;
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
            Assert.AreEqual(expectedResult, Evaluate(formula));
        }

        [TestCase("TRUE & \" to text\"", "TRUE to text")]
        [TestCase("FALSE & \" to text\"", "FALSE to text")]
        [TestCase("true & \" to text\"", "TRUE to text")]
        [TestCase("false & \" to text\"", "FALSE to text")]
        [TestCase("TRUE & FALSE", "TRUEFALSE")]
        public void Concat_ConvertsLogicalToString(string formula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, Evaluate(formula));
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

        [TestCase("#DIV/0! & 1", Error.DivisionByZero)]
        [TestCase("#DIV/0! & \"1\"", Error.DivisionByZero)]
        [TestCase("#REF! & #DIV/0!", Error.CellReference)]
        [TestCase("1 & #NAME?", Error.NameNotRecognized)]
        public void Concat_WithErrorAsOperandReturnsTheError(string formula, Error expectedError)
        {
            Assert.AreEqual(expectedError, Evaluate(formula));
        }

        [Ignore("Arrays are not implemented")]
        [TestCase("{1,2} & \"A\"", "1A")]
        [TestCase("{\"A\",2} & \"B\"", "AB")]
        [TestCase("{TRUE,2} & \"B\"", "TRUEB")]
        [TestCase("{#REF!,5} & 1", Error.CellReference)]
        public void Concat_UsesFirstElementOfArray(string formula, Error expectedError)
        {
            Assert.AreEqual(expectedError, Evaluate(formula));
        }

        #endregion

        #region Unary plus

        [TestCase("+1", 1)]
        [TestCase("+\"1\"", "1")]
        [TestCase("+TRUE", true)]
        [TestCase("+FALSE", false)]
        [TestCase("+#DIV/0!", Error.DivisionByZero)]
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
        [TestCase("-#DIV/0!", Error.DivisionByZero)]
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
        [TestCase("#NAME?%", Error.NameNotRecognized)]
        [TestCase("(1/0)%", Error.DivisionByZero)]
        public void UnaryPercent_ConvertsArgumentBeforePercenting(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Exponentiation

        [TestCase("1^1", 1.0)]
        [TestCase("0^0", Error.NumberInvalid)]
        [TestCase("10^0", 1.0)]
        [TestCase("4^0.5", 2.0)]
        [TestCase("2^0.5", 1.4142135623730951)]
        [TestCase("2^-2", 0.25)]
        [TestCase("\"5\"^\"3\"", 125)]
        [TestCase("5^TRUE", 5)]
        [TestCase("5^FALSE", 1)]
        [TestCase("#VALUE!^1", Error.CellValue)]
        [TestCase("1^#REF!", Error.CellReference)]
        [TestCase("#DIV/0!^#REF!", Error.DivisionByZero)]
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
        [TestCase("#VALUE!*1", Error.CellValue)]
        [TestCase("1*#REF!", Error.CellReference)]
        [TestCase("#DIV/0!*#REF!", Error.DivisionByZero)]
        public void Multiplication_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.That(Evaluate(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
        }

        #endregion

        #region Division

        [TestCase("1/1", 1.0)]
        [TestCase("5/2", 2.5)]
        [TestCase("14.5/2.5", 5.8)]
        [TestCase("10/0", Error.DivisionByZero)]
        [TestCase("0/0", Error.DivisionByZero)]
        [TestCase("2.5/-0.5", -5)]
        [TestCase("\"10\" / \"4\"", 2.5)]
        [TestCase("5/TRUE", 5)]
        [TestCase("5/FALSE", Error.DivisionByZero)]
        [TestCase("#VALUE!/1", Error.CellValue)]
        [TestCase("1/#REF!", Error.CellReference)]
        [TestCase("#DIV/0!/#REF!", Error.DivisionByZero)]
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
        [TestCase("#VALUE! + 1", Error.CellValue)]
        [TestCase("1 + #REF!", Error.CellReference)]
        [TestCase("#DIV/0! + #REF!", Error.DivisionByZero)]
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
        [TestCase("#VALUE! - 1", Error.CellValue)]
        [TestCase("1 - #REF!", Error.CellReference)]
        [TestCase("#DIV/0! - #REF!", Error.DivisionByZero)]
        public void Subtraction_CanWorkWithScalars(string formula, object expectedValue)
        {
            Assert.AreEqual(expectedValue, Evaluate(formula));
        }

        #endregion

        #region Comparison

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
        [TestCase("#VALUE!=#VALUE!", Error.CellValue)]
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
        [TestCase("#VALUE!<>#VALUE!", Error.CellValue)]
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
        [TestCase("#VALUE!>#REF!", Error.CellValue)]
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
        [TestCase("#VALUE!>=#REF!", Error.CellValue)]
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
        [TestCase("#VALUE!<#REF!", Error.CellValue)]
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
        [TestCase("#VALUE!<=#REF!", Error.CellValue)]
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
            Assert.AreEqual(expectedResult, Evaluate(formula));
        }

        #endregion

        // TODO: Replace with XLWorkbook.Evaluate once we switch calculation method.
        private static object Evaluate(string formulaText)
        {
            return XLWorkbook.EvaluateExpr(formulaText);
        }
    }
}
