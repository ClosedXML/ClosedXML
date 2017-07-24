using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML_Tests.Excel
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
                ws.Cell("A1").DataType = XLCellValues.DateTime;

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
    }
}
