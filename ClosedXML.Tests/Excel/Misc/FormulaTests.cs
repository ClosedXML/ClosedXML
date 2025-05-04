using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class FormulaTests
    {
        [Test]
        public void CopyFormula()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").FormulaA1 = "B1";
            ws.Cell("A1").CopyTo("A2");
            Assert.That(ws.Cell("A2").FormulaA1, Is.EqualTo("B2"));
        }

        [Test]
        public void CopyFormula2()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").FormulaA1 = "A2-1";
            ws.Cell("A1").CopyTo("B1");
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").FormulaR1C1, Is.EqualTo("R[1]C-1"));
                Assert.That(ws.Cell("B1").FormulaR1C1, Is.EqualTo("R[1]C-1"));
                Assert.That(ws.Cell("B1").FormulaA1, Is.EqualTo("B2-1"));
            });

            ws.Cell("A1").FormulaA1 = "B1+1";
            ws.Cell("A1").CopyTo("A2");
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").FormulaR1C1, Is.EqualTo("RC[1]+1"));
                Assert.That(ws.Cell("A2").FormulaR1C1, Is.EqualTo("RC[1]+1"));
                Assert.That(ws.Cell("A2").FormulaA1, Is.EqualTo("B2+1"));
            });
        }

        [Test]
        public void CopyFormulaWithSheetNameThatResemblesFormula()
        {
            using var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("S10 Data");
            ws.Cell("A1").Value = "Some value";
            ws.Cell("A2").Value = 123;

            ws = wb.Worksheets.Add("Summary");
            ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
            Assert.That(ws.Cell("A1").Value, Is.EqualTo("Some value"));

            ws.Cell("A1").CopyTo("A2");
            Assert.That(ws.Cell("A2").FormulaA1, Is.EqualTo("'S10 Data'!A2"));

            ws.Cell("A1").CopyTo("B1");
            Assert.That(ws.Cell("B1").FormulaA1, Is.EqualTo("'S10 Data'!B1"));

            ws.Cell("A3").FormulaA1 = "=SUM('S10 Data'!A2)";
            Assert.That(ws.Cell("A3").Value, Is.EqualTo(123));
        }

        [Test]
        public void FormulaWithReferenceIncludingSheetName()
        {
            using var wb = new XLWorkbook();
            object value;
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
            ws.Cell("B1").FormulaA1 = "=SUM(A1:A50)";
            value = ws.Cell("B1").Value;
            Assert.That(value, Is.EqualTo(1275));

            ws = wb.AddWorksheet("Sheet2");

            ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet1!A50)";
            value = ws.Cell("A1").Value;
            Assert.That(value, Is.EqualTo(1275));

            ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:A50)";
            value = ws.Cell("B1").Value;
            Assert.That(value, Is.EqualTo(1275));
        }

        [Test]
        public void InvalidReferences()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
            ws = wb.AddWorksheet("Sheet2");

            ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet2!A50)";
            Assert.That(ws.Cell("A1").Value, Is.EqualTo(XLError.IncompatibleValue));

            ws.Cell("B1").FormulaA1 = "=SUM(UnknownSheet!A50)";
            Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLError.CellReference));
        }

        [Test]
        public void DateAgainstStringComparison()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = new DateTime(2016, 1, 1);

            ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
            var actual = ws.Cell("A2").Value;
            Assert.That("B", Is.EqualTo(actual));

            ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
            actual = ws.Cell("A3").Value;
            Assert.That("B", Is.EqualTo(actual));
        }

        [Test]
        public void FormulaThatReferencesEntireRow()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = 1;
            ws.FirstCell().CellRight().Value = 2;
            ws.FirstCell().CellRight(5).Value = 3;

            ws.FirstCell().CellBelow().FormulaA1 = "=SUM(1:1)";

            var actual = ws.FirstCell().CellBelow().Value;
            Assert.That(actual, Is.EqualTo(6));
        }

        [Test]
        public void FormulaThatReferencesEntireColumn()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().Value = 1;
            ws.FirstCell().CellBelow().Value = 2;
            ws.FirstCell().CellBelow(5).Value = 3;

            ws.FirstCell().CellRight().FormulaA1 = "=SUM(A:A)";

            var actual = ws.FirstCell().CellRight().Value;
            Assert.That(actual, Is.EqualTo(6));
        }

        [Test]
        public void FormulaThatStartsWithEqualsAndPlus()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr("=MID(\"This is a test\", 6, 2)");
            Assert.That(actual, Is.EqualTo("is"));

            actual = XLWorkbook.EvaluateExpr("=+MID(\"This is a test\", 6, 2)");
            Assert.That(actual, Is.EqualTo("is"));

            actual = XLWorkbook.EvaluateExpr("=+++++MID(\"This is a test\", 6, 2)");
            Assert.That(actual, Is.EqualTo("is"));

            actual = XLWorkbook.EvaluateExpr("+MID(\"This is a test\", 6, 2)");
            Assert.That(actual, Is.EqualTo("is"));
        }

        [Test]
        public void UnimplementedStandardFunctionsAreEvaluatedToNameNotFoundError()
        {
            // RTD will never be implemented
            var actual = XLWorkbook.EvaluateExpr("RTD(\"MyRTDServerProdID\",\"MyServer\",\"RaceNum\",\"RunnerID\",\"StatType\")");
            Assert.That(actual, Is.EqualTo(XLError.NameNotRecognized));
        }

        [Test]
        public void FormulasWithErrors()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#REF!)"), Is.EqualTo(XLError.CellReference));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#VALUE!)"), Is.EqualTo(XLError.IncompatibleValue));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#DIV/0!)"), Is.EqualTo(XLError.DivisionByZero));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NAME?)"), Is.EqualTo(XLError.NameNotRecognized));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#N/A)"), Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NULL!)"), Is.EqualTo(XLError.NullValue));
                Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NUM!)"), Is.EqualTo(XLError.NumberInvalid));
            });
        }

        [Test]
        public void LegacyFunctionPropagateErrorWithoutException()
        {
            Assert.That(XLWorkbook.EvaluateExpr("SIN(YEAR(#NAME?))+1"), Is.EqualTo(XLError.NameNotRecognized));
        }

        [Test]
        public void UnicodeLetterParsing()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet C CÄ");
            var ws2 = wb.AddWorksheet("ÖC");
            var ws3 = wb.AddWorksheet("Sheet3");

            ws1.FirstCell().SetValue(100);
            ws2.FirstCell().SetValue(50);

            ws3.FirstCell().FormulaA1 = "='Sheet C CÄ'!A1";
            ws3.FirstCell().CellBelow().FormulaA1 = "ÖC!A1";

            Assert.Multiple(() =>
            {
                Assert.That(ws3.FirstCell().Value, Is.EqualTo(100));
                Assert.That(ws3.FirstCell().CellBelow().Value, Is.EqualTo(50));
            });
        }

        [Test, Ignore("Shifting formulas is done by regexp that breaks array formula.")]
        public void ShiftFormula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B1").FormulaA1 = "ATAN2(C1,C2)";
            ws.Cell("B2").FormulaA1 = "DEC2HEX(C2)";
            ws.Range("B3:B5").FormulaArrayA1 = "DAYS360(C3:C5, D3:D5)";

            ws.Column(1).Delete();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").FormulaA1, Is.EqualTo("ATAN2(B1,B2)"));
                Assert.That(ws.Cell("A2").FormulaA1, Is.EqualTo("DEC2HEX(B2)"));
            });
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A3").HasArrayFormula, Is.True);
                Assert.That(ws.Cell("A3").FormulaA1, Is.EqualTo("DAYS360(B3:B5, C3:C5)"));
            });
        }
    }
}
