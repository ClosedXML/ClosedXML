using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class OperatorsScalarTests
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
        public void ImplicitIntersection_DoesNotAffectSingleCellReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D5").FormulaA1 = "ABS(B3:B3)";

            Assert.AreEqual(1, ws.Cell("D5").Value);
        }

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
        [TestCase("(Sheet1!A1,A5):B5", 10)]
        [TestCase("B5:(Sheet1!A1,A5)", 10)]
        public void Range_UnifiesReferencesIntoSingleAreas(string referenceFormula, int expectedCellCount)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cells("A1:Z100").Value = 1;

            var referenceCells = ws.Evaluate($"SUM({referenceFormula})");
            Assert.AreEqual(expectedCellCount, referenceCells);
        }

        [TestCase("Sheet1!A1:C5")]
        [TestCase("Sheet1!A1:B3:C5")]
        [TestCase("Sheet1!A1:B3:C4:Sheet1!B5:C5")]
        public void Range_LeftSideDeterminesSheetIfRightOmitted(string formula)
        {
            using var wb = new XLWorkbook();
            var firstSheet = wb.AddWorksheet("Sheet1");
            firstSheet.Cells("A1:C5").Value = 1;
            var secondSheet = wb.AddWorksheet("Sheet2");
            secondSheet.Cell("A1").FormulaA1 = $"=SUM({formula})";

            Assert.AreEqual(15, secondSheet.Cell("A1").Value);
        }

        [TestCase("Current!A1:Other!B2")]
        [TestCase("A1:Other!B2")]
        [TestCase("A1:(Other!B2,C3)")]
        [TestCase("Other!A1:(Other!B2,C3)")] // C3 is taken from current worksheet since multiple areas on rhs
        [TestCase("(Other!A1,A5):Other!B2")] // A5 is taken from current worksheet since multiple areas on lhs
        [TestCase("(Current!A1):Other!B2")]
        // [TestCase("Other!A5:(B5)")] This causes #VALUE! in Excel, but it shouldn't. It's likely there is a "Fast parser for simple sheet areas" and "Full path" for complicated operands and they behave inconsistenly
        public void Range_UnificationAcrossSheetsResultsInValueError(string referenceFormula)
        {
            using var wb = new XLWorkbook();
            var formulaSheet = wb.AddWorksheet("Current");
            wb.AddWorksheet("Other");

            // SUM is still legacy, so exception galore!
            Assert.Throws<CellValueException>(() => formulaSheet.Evaluate($"SUM({referenceFormula})"));
        }

        [TestCase("A1:IF(TRUE,1,)")]
        [TestCase("IF(TRUE,1,):A1")]
        [TestCase("IF(TRUE,\"text\"):A1")]
        [TestCase("IF(TRUE,FALSE):A1")]
        public void Range_OnlyReferencesCanBeRange(string referenceFormula)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            // SUM is still legacy, so exception galore!
            Assert.Throws<CellValueException>(() => sheet.Evaluate($"SUM({referenceFormula})"));
        }

        #endregion

        #region Reference union

        [TestCase("A1,A2", 2)]
        [TestCase("A1:A3,B1", 4)]
        [TestCase("A1,B1:B3", 4)]
        [TestCase("Other!A1,Current!A1", 11)]
        [TestCase("A1,Other!A1", 11)]
        [TestCase("B2:D3,B2:D3", 12)] // Full overlap
        [TestCase("A1:B3,B1:C3", 12)] // Partial overlap
        [TestCase("Current!A1:B3,Other!B1:C3", 66)]
        [TestCase("A1,Other!A1,Current!A1", 10 + 1 + 1)]
        [TestCase("A1:B2,Other!A1:B2,B2:C3,Other!E5:Other!F6", 4 + 40 + 4 + 40)]
        public void Union_CanJoinAnyTwoRanges(string formula, int expectedSum)
        {
            using var wb = new XLWorkbook();
            var currentSheet = wb.AddWorksheet("Current");
            currentSheet.Cells("A1:F10").Value = 1;
            var otherSheet = wb.AddWorksheet("Other");
            otherSheet.Cells("A1:F10").Value = 10;

            // Not extra braces, so the comma is interpreted as union and not an extra argument
            var value = currentSheet.Evaluate($"SUM(({formula}))");

            Assert.AreEqual(expectedSum, value);
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
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
            Assert.That(XLWorkbook.EvaluateExpr(formula), Is.EqualTo(expectedValue).Within(XLHelper.Epsilon));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedValue, XLWorkbook.EvaluateExpr(formula));
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
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("\"\">10", true)]
        [TestCase("\"1\">10", true)]
        [TestCase("10<\"\"", true)]
        [TestCase("10<\"1\"", true)]
        public void Comparison_TextIsAlwaysGreaterThanAnyNumber(string formula, bool expectedResult)
        {
            Assert.AreEqual(expectedResult, XLWorkbook.EvaluateExpr(formula));
        }

        #endregion
    }
}
