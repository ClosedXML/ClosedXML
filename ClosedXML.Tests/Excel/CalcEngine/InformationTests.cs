using ClosedXML.Excel;
using NUnit.Framework;
using System;

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
            Assert.That(ws.Evaluate($"ERROR.TYPE({argumentFormula})"), Is.EqualTo(XLError.NoValueAvailable));
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
            Assert.That(XLWorkbook.EvaluateExpr($"ERROR.TYPE({error})"), Is.EqualTo(expectedNumber));
        }

        #region IsBlank Tests

        [Test]
        public void IsBlank_EmptyCell_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate("IsBlank(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        [Test]
        public void IsBlank_NonEmptyCell_False()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "1";
            var actual = ws.Evaluate("IsBlank(A1)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("5")]
        [TestCase("\"\"")]
        [TestCase("\"Hello\"")]
        [TestCase("#DIV/0!")]
        public void IsBlank_NonEmptyValue_False(string value)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsBlank({value})");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsBlank_InlineBlank_True()
        {
            var actual = XLWorkbook.EvaluateExpr("IsBlank(IF(TRUE,,))");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsBlank Tests

        [TestCase("IF(TRUE,,)")]
        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("\"text\"")]
        public void IsErr_NonErrorValues_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
            Assert.That(actual, Is.EqualTo(false));
        }

        [TestCase("#DIV/0!")]
        [TestCase("#NAME?")]
        [TestCase("#NULL!")]
        [TestCase("#NUM!")]
        [TestCase("#REF!")]
        [TestCase("#VALUE!")]
        public void IsErr_ErrorsExceptNA_True(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
            Assert.That(actual, Is.EqualTo(true));
        }

        [Test]
        public void IsErr_NA_False()
        {
            var actual = XLWorkbook.EvaluateExpr("IsErr(#N/A)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [TestCase("#DIV/0!")]
        [TestCase("#N/A")]
        [TestCase("#NAME?")]
        [TestCase("#NULL!")]
        [TestCase("#NUM!")]
        [TestCase("#REF!")]
        [TestCase("#VALUE!")]
        public void IsError_Errors_True(string error)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsError({error})");
            Assert.That(actual, Is.EqualTo(true));
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("FALSE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("\"text\"")]
        public void IsError_NonErrors_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsError({valueFormula})");
            Assert.That(actual, Is.EqualTo(false));
        }

        #region IsEven Tests

        [TestCase("2")]
        [TestCase("\"1 2/2\"")]
        [TestCase("\"4 1/2\"")]
        [TestCase("\"48:30:00\"")]
        [TestCase("\"1900-01-02\"")]
        public void IsEven_NumberLikeValue_ConvertedThroughValueSemantic(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsEven({valueFormula})");
            Assert.That(actual, Is.EqualTo(true));
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
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsEven(A4)");
            Assert.That(actual, Is.EqualTo(true));
        }

        [Test]
        [Ignore("Arrays not yet implemented.")]
        public void IsEven_Array_ReturnsArray()
        {
            Assert.That(XLWorkbook.EvaluateExpr("SUM(N(IsEven({\"2.9\";2;1})))"), Is.EqualTo(2.0));
        }

        [Test]
        public void IsEven_ReferenceToMoreThanOneCell_Error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 2).FormulaA1 = "IsEven(A1:A2)";
            Assert.That(ws.Cell(1, 2).Value, Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("TRUE", XLError.IncompatibleValue)]
        [TestCase("FALSE", XLError.IncompatibleValue)]
        [TestCase("\"\"", XLError.IncompatibleValue)]
        [TestCase("\"test\"", XLError.IncompatibleValue)]
        [TestCase("#DIV/0!", XLError.DivisionByZero)]
        [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
        public void IsEven_NonNumberValues_Error(string valueFormula, XLError expectedError)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"IsEven({valueFormula})"), Is.EqualTo(expectedError));
        }

        #endregion IsEven Tests

        #region IsLogical Tests

        [TestCase("TRUE")]
        [TestCase("FALSE")]
        public void IsLogical_OnlyLogical_True(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
            Assert.That(actual, Is.EqualTo(true));
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
        public void IsLogical_NonLogicalValue_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsLogical_ReferenceToLogicalValue_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = true;

            var actual = ws.Evaluate("IsLogical(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsLogical Tests

        [Test]
        public void IsNA_NA_True()
        {
            var actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
            Assert.That(actual, Is.EqualTo(true));
        }

        [TestCase("IF(TRUE,,)")]
        [TestCase("TRUE")]
        [TestCase("0")]
        [TestCase("\"\"")]
        [TestCase("#REF!")]
        [TestCase("\"#N/A\"")]
        public void IsNA_NonNotAvailableValue_False(string valueFormula)
        {
            var actual = XLWorkbook.EvaluateExpr($"ISNA({valueFormula})");
            Assert.That(actual, Is.EqualTo(false));
        }

        #region IsNotText Tests

        [Test]
        public void IsNotText_ReferenceToBlankCell_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate("IsNonText(A1)");
            Assert.That(actual, Is.EqualTo(true));
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
            Assert.That(actual, Is.EqualTo(false));
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
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("IsNonText(A2)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("IsNonText(A3)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("IsNonText(A4)");
            Assert.That(actual, Is.EqualTo(true));
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
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("IsNumber(A2)");
            Assert.That(actual, Is.EqualTo(false));
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
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A3)");
            Assert.That(actual, Is.EqualTo(true));
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
            Assert.That(actual, Is.EqualTo(false));
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
            Assert.That(actual, Is.EqualTo(true));
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
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsOdd(A4)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [SetCulture("en-US")]
        [Test]
        [Ignore("Arrays not yet implemented.")]
        public void IsOdd_Array_ReturnsArray()
        {
            Assert.That(XLWorkbook.EvaluateExpr("SUM(N(IsOdd({\"3.2\",7,2})))"), Is.EqualTo(2.0));
        }

        [Test]
        public void IsOdd_ReferenceToMoreThanOneCell_Error()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 2).FormulaA1 = "IsOdd(A1:A2)";
            Assert.That(ws.Cell(1, 2).Value, Is.EqualTo(XLError.IncompatibleValue));
        }

        [TestCase("TRUE", XLError.IncompatibleValue)]
        [TestCase("FALSE", XLError.IncompatibleValue)]
        [TestCase("\"\"", XLError.IncompatibleValue)]
        [TestCase("\"test\"", XLError.IncompatibleValue)]
        [TestCase("#DIV/0!", XLError.DivisionByZero)]
        [TestCase("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
        public void IsOdd_NonNumberValues_Error(string valueFormula, XLError expectedError)
        {
            Assert.That(XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})"), Is.EqualTo(expectedError));
        }

        #endregion IsOdd Test

        [TestCase("A1")]
        [TestCase("(A1,A5)")]
        public void IsRef_Reference_True(string reference)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123";

            ws.Cell("B1").FormulaA1 = $"ISREF({reference})";

            Assert.That(ws.Cell("B1").Value, Is.EqualTo(true));
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

            Assert.That(ws.Cell("B1").Value, Is.EqualTo(false));
        }

        #region IsText Tests

        [Test]
        public void IsText_BlankCell_False()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B1").FormulaA1 = "ISTEXT(A1)";

            Assert.That(ws.Cell("B1").Value, Is.EqualTo(false));
        }

        [TestCase("0")]
        [TestCase("123")]
        [TestCase("TRUE")]
        [TestCase("#DIV/0!")]
        [TestCase("IF(TRUE,,)")]
        public void IsText_NonText_False(string nonText)
        {
            var actual = XLWorkbook.EvaluateExpr($"ISTEXT({nonText})");
            Assert.That(actual, Is.EqualTo(false));
        }

        [TestCase("")]
        [TestCase("abc")]
        public void IsText_CellWithText_True(string textValue)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = textValue;

            var actual = ws.Evaluate("IsText(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsText Tests

        #region N Tests

        [Test]
        public void N_Blank_Zero()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(0.0));
        }

        [Test]
        public void N_Date_SerialNumber()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var testedDate = DateTime.Now;
            ws.Cell("A1").Value = testedDate;
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(testedDate.ToSerialDateTime()));
        }

        [Test]
        public void N_False_Zero()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = false;
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(0));
        }

        [Test]
        public void N_True_One()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = true;
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(1));
        }
        [Test]
        public void N_Number_Number()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var testedValue = 123;
            ws.Cell("A1").Value = testedValue;
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(testedValue));
        }

        [TestCase("")]
        [TestCase("abc")]
        public void N_String_Zero(string text)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = text;
            var actual = ws.Evaluate("N(A1)");
            Assert.That(actual, Is.EqualTo(0));
        }

        [Test]
        [Ignore("Array not implemented")]
        public void N_Array_ConvertsIndividualItems()
        {
            var actual = XLWorkbook.EvaluateExpr("SUM(N({2,TRUE}))");
            Assert.That(actual, Is.EqualTo(3));
        }

        [TestCase("A1")]
        [TestCase("A1:B1")]
        [TestCase("(A1, B1)")]
        public void N_Reference_TakesFirstCellFromFirstArea(string reference)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 5;
            ws.Cell("B1").Value = 10;

            var actual = ws.Evaluate($"SUM(N({reference}))");
            Assert.That(actual, Is.EqualTo(5));
        }

        #endregion N Tests

        [TestCase("IF(TRUE,,)", 1)]
        [TestCase("0", 1)]
        [TestCase("1", 1)]
        [TestCase("-5.2", 1)]
        [TestCase("\"\"", 2)]
        [TestCase("\"text\"", 2)]
        [TestCase("\"1\"", 2)]
        [TestCase("\"TRUE\"", 2)]
        [TestCase("TRUE", 4)]
        [TestCase("FALSE", 4)]
        [TestCase("#DIV/0!", 16)]
        [TestCase("1/0", 16)]
        [TestCase("#N/A", 16)]
        [TestCase("#VALUE!", 16)]
        public void Type_NonReferenceScalarValues(string literalValues, double expectedNumber)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = $"TYPE({literalValues})";
            Assert.That(ws.Cell("A1").Value, Is.EqualTo(expectedNumber));
        }

        [Ignore("Arrays not implemented")]
        [TestCase("{1}")]
        [TestCase("{TRUE,#N/A}")]
        [TestCase("{\"abc\";5}")]
        public void Type_Array_HasValue64(string arrayLiteral)
        {
            var actual = XLWorkbook.EvaluateExpr($"TYPE({arrayLiteral})");
            Assert.That(actual, Is.EqualTo(64.0));
        }

        [TestCase("A1:A2")]
        // [TestCase("(A1:A3 A2:B3)")] Not implemented // Intersection results in a 1x2 block
        public void Type_ReferenceToNonSingleCell_BehavesLikeArray(string reference)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("C1").FormulaA1 = $"TYPE({reference})";
            Assert.That(ws.Cell("C1").Value, Is.EqualTo(64.0));
        }

        [Test]
        public void Type_ReferenceToSingleCell_ReturnsTypeOfCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "text";

            ws.Cell("C1").FormulaA1 = "TYPE(A1)";
            Assert.That(ws.Cell("C1").Value, Is.EqualTo(2.0));
        }

        [Test]
        public void Type_MultiAreaReference_ReturnsError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "text";

            ws.Cell("C1").FormulaA1 = "TYPE((A1,A1))";
            Assert.That(ws.Cell("C1").Value, Is.EqualTo(16.0));
        }
    }
}
