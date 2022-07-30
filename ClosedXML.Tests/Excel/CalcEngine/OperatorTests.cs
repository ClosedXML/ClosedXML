using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

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

        [Test]
        public void ImplicitIntersection_TakesReferenceFromVerticalLine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("B5").FormulaA1 = "ABS(A3:Z3)";

            Assert.AreEqual(1, ws.Cell("B5").Value);
        }

        [Test]
        public void ImplicitIntersection_TakesReferenceEvenFromIntersectionEvenFromDifferentSheet()
        {
            using var wb = new XLWorkbook();
            var sheet1 = wb.AddWorksheet("Sheet1");
            sheet1.Cell("B3").Value = -1;

            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet2.Cell("D3").FormulaA1 = "ABS(Sheet1!B1:B10)";

            Assert.AreEqual(1, sheet2.Cell("D3").Value);
        }

        [Test]
        public void ImplicitIntersection_WithoutIntersectionResultsInValueError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D5").FormulaA1 = "ABS(B1:B4)";

            Assert.AreEqual(Error.CellValue, ws.Cell("D5").Value);
        }

        [Test]
        public void ImplicitIntersection_CanWorkOnlyWithOneArea()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D3").FormulaA1 = "ABS((B1:B2,B3:B5))"; // A continous range made of two areas

            Assert.AreEqual(Error.CellValue, ws.Cell("D3").Value);
        }

        [Test]
        public void ImplicitIntersection_IntersectionMustHaveSpanOfOneCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            var horizontalIntersectionCell = ws.Cell("D3");
            horizontalIntersectionCell.FormulaA1 = "ABS(A1:B5)";
            Assert.AreEqual(Error.CellValue, horizontalIntersectionCell.Value);

            var verticalIntersectionCell = ws.Cell("B5");
            verticalIntersectionCell.FormulaA1 = "ABS(A3:C4)";
            Assert.AreEqual(Error.CellValue, verticalIntersectionCell.Value);
        }

        #endregion

        #region Reference range operator

        [TestCase("A1:B2", 4)]
        [TestCase("A1:B5:C3", 3 * 5)]
        [TestCase("A1:C3:B5", 3 * 5)]
        [TestCase("A1:C3:B2", 3 * 3)]
        [TestCase("Sheet1!A1:B2", 4)]
        [TestCase("Sheet1!A1:Sheet1!B2", 4)]
        [TestCase("Sheet1!A1:Sheet1!B2", 4)]
        [TestCase("A1:Sheet1!B2", 4)]
        [TestCase("Sheet1!B2:C5:Sheet1!D3", 12)]
        public void Range_UnifiesReferencesIntoSingleAreas(string referenceFormula, int expectedCellCount)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cells("A1:Z100").Value = 1;

            var referenceCells = ws.Evaluate($"SUM({referenceFormula})");
            Assert.AreEqual(expectedCellCount, referenceCells);
        }

        #endregion
    }
}
