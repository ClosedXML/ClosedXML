using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class OperatorTests
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
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("#DIV/0! & 1", Error.DivisionByZero)]
        [TestCase("#DIV/0! & \"1\"", Error.DivisionByZero)]
        [TestCase("#REF! & #DIV/0!", Error.CellReference)]
        [TestCase("1 & #NAME?", Error.NameNotRecognized)]
        public void Concat_WithErrorAsOperandReturnsTheError(string formula, Error expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr(formula));
        }

        [Ignore("Arrays are not implemented")]
        [TestCase("{1,2} & \"A\"", "1A")]
        [TestCase("{\"A\",2} & \"B\"", "AB")]
        [TestCase("{TRUE,2} & \"B\"", "TRUEB")]
        [TestCase("{#REF!,5} & 1", Error.CellReference)]
        public void Concat_UsesFirstElementOfArray(string formula, Error expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr(formula));
        }

        #endregion

        #region Implicit intersection

        [Test]
        public void ImplicitIntersection_TakesReferenceFromHorizontalLine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D3").FormulaA1 = "ABS(B1:B10)";

            Assert.AreEqual(1, ws.Cell("D3").Value);
        }

        #endregion
    }
}
