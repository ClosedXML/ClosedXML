using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
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
            Assert.AreEqual("B2", ws.Cell("A2").FormulaA1);
        }

        [Test]
        public void CopyFormula2()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").FormulaA1 = "A2-1";
                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("R[1]C-1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("R[1]C-1", ws.Cell("B1").FormulaR1C1);
                Assert.AreEqual("B2-1", ws.Cell("B1").FormulaA1);

                ws.Cell("A1").FormulaA1 = "B1+1";
                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("RC[1]+1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("RC[1]+1", ws.Cell("A2").FormulaR1C1);
                Assert.AreEqual("B2+1", ws.Cell("A2").FormulaA1);
            }
        }

        [Test]
        public void CopyFormulaWithSheetNameThatResemblesFormula()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("S10 Data");
                ws.Cell("A1").Value = "Some value";
                ws.Cell("A2").Value = 123;

                ws = wb.Worksheets.Add("Summary");
                ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
                Assert.AreEqual("Some value", ws.Cell("A1").Value);

                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("'S10 Data'!A2", ws.Cell("A2").FormulaA1);

                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("'S10 Data'!B1", ws.Cell("B1").FormulaA1);

                ws.Cell("A3").FormulaA1 = "=SUM('S10 Data'!A2)";
                Assert.AreEqual(123, ws.Cell("A3").Value);
            }
        }

        [Test]
        public void FormulaWithReferenceIncludingSheetName()
        {
            using (var wb = new XLWorkbook())
            {
                object value;
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
                ws.Cell("B1").FormulaA1 = "=SUM(A1:A50)";
                value = ws.Cell("B1").Value;
                Assert.AreEqual(1275, value);

                ws = wb.AddWorksheet("Sheet2");

                ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet1!A50)";
                value = ws.Cell("A1").Value;
                Assert.AreEqual(1275, value);

                ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:A50)";
                value = ws.Cell("B1").Value;
                Assert.AreEqual(1275, value);
            }
        }

        [Test]
        public void InvalidReferences()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
                ws = wb.AddWorksheet("Sheet2");

                ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet2!A50)";
                Assert.That(() => ws.Cell("A1").Value, Throws.InstanceOf<ArgumentOutOfRangeException>());

                ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:UnknownSheet!A50)";
                Assert.That(() => ws.Cell("B1").Value, Throws.InstanceOf<ArgumentOutOfRangeException>());
            }
        }

        [Test]
        public void DateAgainstStringComparison()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = new DateTime(2016, 1, 1);
                ws.Cell("A1").DataType = XLDataType.DateTime;

                ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
                var actual = ws.Cell("A2").Value;
                Assert.AreEqual(actual, "B");

                ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
                actual = ws.Cell("A3").Value;
                Assert.AreEqual(actual, "B");
            }
        }

        [Test]
        public void FormulaThatReferencesEntireRow()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().Value = 1;
                ws.FirstCell().CellRight().Value = 2;
                ws.FirstCell().CellRight(5).Value = 3;

                ws.FirstCell().CellBelow().FormulaA1 = "=SUM(1:1)";

                var actual = ws.FirstCell().CellBelow().Value;
                Assert.AreEqual(6, actual);
            }
        }

        [Test]
        public void FormulaThatReferencesEntireColumn()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().Value = 1;
                ws.FirstCell().CellBelow().Value = 2;
                ws.FirstCell().CellBelow(5).Value = 3;

                ws.FirstCell().CellRight().FormulaA1 = "=SUM(A:A)";

                var actual = ws.FirstCell().CellRight().Value;
                Assert.AreEqual(6, actual);
            }
        }

        [Test]
        public void FormulaThatStartsWithEqualsAndPlus()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr("=MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("=+MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("=+++++MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("+MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);
        }

        [Test]
        public void FormulasWithErrors()
        {
            Assert.AreEqual(XLCalculationErrorType.CellReference, XLWorkbook.EvaluateExpr("YEAR(#REF!)"));
            Assert.AreEqual(XLCalculationErrorType.CellValue, XLWorkbook.EvaluateExpr("YEAR(#VALUE!)"));
            Assert.AreEqual(XLCalculationErrorType.DivisionByZero, XLWorkbook.EvaluateExpr("YEAR(#DIV/0!)"));
            Assert.AreEqual(XLCalculationErrorType.NameNotRecognized, XLWorkbook.EvaluateExpr("YEAR(#NAME?)"));
            Assert.AreEqual(XLCalculationErrorType.NoValueAvailable, XLWorkbook.EvaluateExpr("YEAR(#N/A)"));
            Assert.AreEqual(XLCalculationErrorType.NullValue, XLWorkbook.EvaluateExpr("YEAR(#NULL!)"));
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr("YEAR(#NUM!)"));
        }

        [Test]
        public void UnicodeLetterParsing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet C CÄ");
                var ws2 = wb.AddWorksheet("ÖC");
                var ws3 = wb.AddWorksheet("Sheet3");

                ws1.FirstCell().SetValue(100);
                ws2.FirstCell().SetValue(50);

                ws3.FirstCell().FormulaA1 = "='Sheet C CÄ'!A1";
                ws3.FirstCell().CellBelow().FormulaA1 = "ÖC!A1";

                Assert.AreEqual(100, ws3.FirstCell().Value);
                Assert.AreEqual(50, ws3.FirstCell().CellBelow().Value);
            }
        }

        [Test]
        public void ShiftFormula()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.Cell("B1").FormulaA1 = "ATAN2(C1,C2)";
                ws.Cell("B2").FormulaA1 = "DEC2HEX(C2)";
                ws.Range("B3:B5").FormulaA1 = "{DAYS360(C3:C5, D3:D5)}";

                ws.Column(1).Delete();

                Assert.AreEqual("ATAN2(B1,B2)", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("DEC2HEX(B2)", ws.Cell("A2").FormulaA1);
                Assert.AreEqual("{DAYS360(B3:B5, C3:C5)}", ws.Cell("A3").FormulaA1);
            }
        }
    }
}
