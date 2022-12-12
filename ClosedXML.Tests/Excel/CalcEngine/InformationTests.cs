using ClosedXML.Excel;
using NUnit.Framework;
using System;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class InformationTests
    {
        [TestCase("A1")] // blank
        [TestCase("TRUE")]
        [TestCase("14.5")]
        [TestCase("\"text\"")]
        public void ErrorType_NonErrorsAreNA(string argumentFormula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate($"ERROR.TYPE({argumentFormula})"));
        }

        [TestCase("#NULL!", 1)]
        [TestCase("#DIV/0!", 2)]
        [TestCase("#VALUE!", 3)]
        [TestCase("#REF!", 4)]
        [TestCase("#NAME?", 5)]
        [TestCase("#NUM!", 6)]
        [TestCase("#N/A", 7)]
        //[TestCase("#GETTING_DATA", 8)] OLAP Cube not supported
        public void ErrorType_ReturnsNumberForError(string error, int expectedNumber)
        {
            Assert.AreEqual(expectedNumber, XLWorkbook.EvaluateExpr($"ERROR.TYPE({error})"));
        }

        #region IsBlank Tests

        [Test]
        public void IsBlank_EmptyCell_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate("IsBlank(A1)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsBlank_NonEmptyCell_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "1";
            var actual = ws.Evaluate("IsBlank(A1)");
            Assert.AreEqual(false, actual);
        }

        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("5")]
        [TestCase("\"\"")]
        [TestCase("\"Hello\"")]
        [TestCase("#DIV/0!")]
        public void IsBlank_NonEmptyValue_false(string value)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsBlank({value})");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsBlank_InlineBlank_true()
        {
            var actual = XLWorkbook.EvaluateExpr("IsBlank(IF(TRUE,,))");
            Assert.AreEqual(true, actual);
        }

        #endregion IsBlank Tests

        [TestCase("IF(TRUE,,)")]
        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("\"text\"")]
        public void IsErr_NonErrorValues_false(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
            Assert.AreEqual(false, actual);
        }

        [TestCase("#DIV/0!")]
        [TestCase("#NAME?")]
        [TestCase("#NULL!")]
        [TestCase("#NUM!")]
        [TestCase("#REF!")]
        [TestCase("#VALUE!")]
        public void IsErr_ErrorsExceptNA_true(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsErr_NA_false()
        {
            var actual = XLWorkbook.EvaluateExpr("IsErr(#N/A)");
            Assert.AreEqual(false, actual);
        }

        [TestCase("#DIV/0!")]
        [TestCase("#N/A")]
        [TestCase("#NAME?")]
        [TestCase("#NULL!")]
        [TestCase("#NUM!")]
        [TestCase("#REF!")]
        [TestCase("#VALUE!")]
        public void IsError_Errors_true(string error)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsError({error})");
            Assert.AreEqual(true, actual);
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("\"text\"")]
        public void IsError_NonErrors_false(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsError({valueFormula})");
            Assert.AreEqual(false, actual);
        }

        #region IsEven Tests

        [SetCulture("en-US")]
        [TestCase("2")]
        [TestCase("\"1 2/2\"")]
        [TestCase("\"4 1/2\"")]
        [TestCase("\"48:30:00\"")]
        [TestCase("\"1900-01-02\"")]
        public void IsEven_SingleValue_ConvertedThroughValueSemantic(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsEven({valueFormula})");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsEven_NonIntegerValues_TruncatedForEvaluation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 4;
            ws.Cell("A2").Value = 0.9;
            ws.Cell("A3").Value = -2.9;

            var actual = ws.Evaluate("=IsEven(A1)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsEven(A4)");
            Assert.AreEqual(true, actual);
        }

        [SetCulture("en-US")]
        [Test]
        [Ignore("Arrays not yet implemented.")]
        public void IsEven_Array_ReturnsArray()
        {
            Assert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsEven({\"2.9\";2;1})))"));
        }

        [Test]
        public void IsEven_ReferenceToMoreThanOneCell_Error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 2).FormulaA1 = "IsEven(A1:A2)";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
        }

        [TestCase("TRUE", XLError.IncompatibleValue)]
        [TestCase("FALSE", XLError.IncompatibleValue)]
        [TestCase("\"\"", XLError.IncompatibleValue)]
        [TestCase("\"test\"", XLError.IncompatibleValue)]
        [TestCase("#DIV/0!", XLError.DivisionByZero)]
        [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
        public void IsEven_NonNumberValues_Error(string valueFormula, XLError expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsEven({valueFormula})"));
        }

        #endregion IsEven Tests

        #region IsLogical Tests

        [TestCase("TRUE")]
        [TestCase("FALSE")]
        public void IsLogical_OnlyLogical_True(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
            Assert.AreEqual(true, actual);
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("0")]
        [TestCase("1")]
        [TestCase("\"\"")]
        [TestCase("\"text\"")]
        [TestCase("#NAME?")]
        [TestCase("#N/A")]
        [TestCase("#VALUE!")]
        [TestCase("#REF!")]
        public void IsLogical_NonLogical_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsLogical_ReferenceToLogical_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = true;

            var actual = ws.Evaluate("IsLogical(A1)");
            Assert.AreEqual(true, actual);
        }

        #endregion IsLogical Tests

        [Test]
        public void IsNA_NA_True()
        {
            var actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
            Assert.AreEqual(true, actual);
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("TRUE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("#REF!")]
        [TestCase("\"#N/A\"")]
        public void IsNA_NA_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"ISNA({valueFormula})");
            Assert.AreEqual(false, actual);
        }

        #region IsNotText Tests

        [Test]
        public void IsNotText_ReferenceToBlankCell_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate("IsNonText(A1)");
            Assert.AreEqual(true, actual);
        }

        [TestCase("")]
        [TestCase("  ")]
        [TestCase("text")]
        public void IsNotText_ReferenceToStringCell_False(string text)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = text;
            var actual = ws.Evaluate("IsNonText(A1)");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsNotText_NonTextValues_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = 123; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = true; //Bool Value
            ws.Cell("A4").Value = XLError.IncompatibleValue; //Error value

            var actual = ws.Evaluate("IsNonText(A1)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("IsNonText(A2)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("IsNonText(A3)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("IsNonText(A4)");
            Assert.AreEqual(true, actual);
        }

        #endregion IsNotText Tests

        #region IsNumber Tests

        [Test]
        public void IsNumber_Simple_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd"; //String Value
            ws.Cell("A2").Value = true; //Bool Value

            var actual = ws.Evaluate("IsNumber(A1)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("IsNumber(A2)");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsNumber_Simple_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = 123; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = new TimeSpan(2, 30, 50); //TimeSpan Value

            var actual = ws.Evaluate("=IsNumber(A1)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A3)");
            Assert.AreEqual(true, actual);
        }

        [TestCase("TRUE")]
        [TestCase("FALSE")]
        [TestCase("\"\"")]
        [TestCase("#DIV/0!")]
        [TestCase("#NULL!")]
        [TestCase("#VALUE!")]
        [TestCase("#N/A")]
        public void IsNumber_NonNumber_False(string nonNumberValue)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsNumber({nonNumberValue})");
            Assert.AreEqual(false, actual);
        }

        #endregion IsNumber Tests

        #region IsOdd Test

        [SetCulture("en-US")]
        [TestCase("1")]
        [TestCase("\"2 3/3\"")]
        [TestCase("\"5 1/3\"")]
        [TestCase("\"25:30:00\"")]
        [TestCase("\"1900-01-03\"")]
        public void IsOdd_SingleValue_ConvertedThroughValueSemantic(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsOdd_NonIntegerValues_TruncatedForEvaluation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 3;
            ws.Cell("A2").Value = 1.9;
            ws.Cell("A3").Value = -5.9;

            var actual = ws.Evaluate("=IsOdd(A1)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsOdd(A4)");
            Assert.AreEqual(false, actual);
        }

        [SetCulture("en-US")]
        [Test]
        [Ignore("Arrays not yet implemented.")]
        public void IsOdd_Array_ReturnsArray()
        {
            Assert.AreEqual(2.0, XLWorkbook.EvaluateExpr("SUM(N(IsOdd({\"3.2\",7,2})))"));
        }

        [Test]
        public void IsOdd_ReferenceToMoreThanOneCell_Error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 2).FormulaA1 = "IsOdd(A1:A2)";
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell(1, 2).Value);
        }

        [TestCase("TRUE", XLError.IncompatibleValue)]
        [TestCase("FALSE", XLError.IncompatibleValue)]
        [TestCase("\"\"", XLError.IncompatibleValue)]
        [TestCase("\"test\"", XLError.IncompatibleValue)]
        [TestCase("#DIV/0!", XLError.DivisionByZero)]
        [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
        public void IsOdd_NonNumberValues_Error(string valueFormula, XLError expectedError)
        {
            Assert.AreEqual(expectedError, XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})"));
        }

        #endregion IsOdd Test

        [TestCase("A1")]
        [TestCase("(A1,A5)")]
        public void IsRef_True(string reference)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123";

            ws.Cell("B1").FormulaA1 = $"ISREF({reference})";

            Assert.AreEqual(true, ws.Cell("B1").Value);
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("TRUE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        // [TestCase("{1;2}")] Arrays not yet implemented
        [TestCase("#N/A")]
        [TestCase("#VALUE!")]
        public void IsRef_NonReference_False(string nonReference)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("B1").FormulaA1 = $"ISREF({nonReference})";

            Assert.AreEqual(false, ws.Cell("B1").Value);
        }

        #region IsText Tests

        [Test]
        public void IsText_BlankCell_False()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B1").FormulaA1 = "ISTEXT(A1)";

            Assert.AreEqual(false, ws.Cell("B1").Value);
        }

        [TestCase("0")]
        [TestCase("123")]
        [TestCase("TRUE")]
        [TestCase("#DIV/0!")]
        [TestCase("IF(TRUE,,)")]
        public void IsText_NonText_False(string nonText)
        {
            var actual = XLWorkbook.EvaluateExpr($"ISTEXT({nonText})");
            Assert.AreEqual(false, actual);
        }

        [TestCase("")]
        [TestCase("abc")]
        public void IsText_CellWithText_True(string textValue)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = textValue;

            var actual = ws.Evaluate("IsText(A1)");
            Assert.AreEqual(true, actual);
        }

        #endregion IsText Tests

        #region N Tests

        [Test]
        public void N_Date_SerialNumber()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet");
                var testedDate = DateTime.Now;
                ws.Cell("A1").Value = testedDate;
                var actual = ws.Evaluate("=N(A1)");
                Assert.AreEqual(testedDate.ToOADate(), actual);
            }
        }

        [Test]
        public void N_False_Zero()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet");
                ws.Cell("A1").Value = false;
                var actual = ws.Evaluate("=N(A1)");
                Assert.AreEqual(0, actual);
            }
        }

        [Test]
        public void N_Number_Number()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet");
                var testedValue = 123;
                ws.Cell("A1").Value = testedValue;
                var actual = ws.Evaluate("=N(A1)");
                Assert.AreEqual(testedValue, actual);
            }
        }

        [Test]
        public void N_String_Zero()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet");
                ws.Cell("A1").Value = "asd";
                var actual = ws.Evaluate("=N(A1)");
                Assert.AreEqual(0, actual);
            }
        }

        [Test]
        public void N_True_One()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet");
                ws.Cell("A1").Value = true;
                var actual = ws.Evaluate("=N(A1)");
                Assert.AreEqual(1, actual);
            }
        }

        #endregion N Tests
    }
}
