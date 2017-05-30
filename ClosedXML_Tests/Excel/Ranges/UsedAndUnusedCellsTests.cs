using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML_Tests.Excel.Ranges
{
    [TestFixture]
    public class UsedAndUnusedCellsTests
    {
        private XLWorkbook workbook;

        [OneTimeSetUp]
        public void SetupWorkbook()
        {
            workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");
            ws.Cell(1, 1).Value = "A1";
            ws.Cell(1, 3).Value = "C1";
            ws.Cell(2, 2).Value = "B2";
            ws.Cell(4, 1).Value = "A4";
            ws.Cell(5, 2).Value = "B5";
        }

        [Test]
        public void CountUsedCellsInRow()
        {
            int i = 0;
            var row = workbook.Worksheets.First().FirstRow();
            foreach (var cell in row.Cells()) // Cells() returns UnUsed cells by default
            {
                i++;
            }
            Assert.AreEqual(2, i);

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow();
            foreach (var cell in row.Cells())
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountAllCellsInRow()
        {
            int i = 0;
            var row = workbook.Worksheets.First().FirstRow();
            foreach (var cell in row.Cells(false)) // All cells in range between first and last cells used
            {
                i++;
            }
            Assert.AreEqual(3, i);

            i = 0;
            row = workbook.Worksheets.First().FirstRow().RowBelow(); //This row has no empty cells BETWEEN used cells
            foreach (var cell in row.Cells(false))
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountUsedCellsInColumn()
        {
            int i = 0;
            var column = workbook.Worksheets.First().FirstColumn();
            foreach (var cell in column.Cells()) // Cells() returns UnUsed cells by default
            {
                i++;
            }
            Assert.AreEqual(2, i);

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight();
            foreach (var cell in column.Cells())
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountAllCellsInColumn()
        {
            int i = 0;
            var column = workbook.Worksheets.First().FirstColumn();
            foreach (var cell in column.Cells(false)) // All cells in range between first and last cells used
            {
                i++;
            }
            Assert.AreEqual(4, i);

            i = 0;
            column = workbook.Worksheets.First().FirstColumn().ColumnRight().ColumnRight(); //This column has no empty cells BETWEEN used cells
            foreach (var cell in column.Cells(false))
            {
                i++;
            }
            Assert.AreEqual(1, i);
        }

        [Test]
        public void CountUsedCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.Cells()) // Only used cells in worksheet
            {
                i++;
            }
            Assert.AreEqual(5, i);
        }

        [Test]
        public void CountAllCellsInWorksheet()
        {
            var ws = workbook.Worksheets.First();
            int i = 0;

            foreach (var cell in ws.Cells(false)) // All cells in range between first and last cells used (cartesian product of range)
            {
                i++;
            }
            Assert.AreEqual(15, i);
        }
    }
}
