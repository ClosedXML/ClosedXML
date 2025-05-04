using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class ValueSliceTests
    {
        [Test]
        public void Deleting_worksheet_dereferences_all_texts_in_its_value_slice()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var keptWs = wb.AddWorksheet();
            var removedWs = wb.AddWorksheet();
            keptWs.Cell("A1").Value = "Double referenced text";
            removedWs.Cell("A1").Value = "Double referenced text";
            removedWs.Cell("B1").Value = "Single referenced text";

            Assert.That(sst.Count, Is.EqualTo(2));

            wb.Worksheets.Delete(removedWs.Name);

            Assert.Multiple(() =>
            {
                Assert.That(sst.Count, Is.EqualTo(1));
                Assert.That(keptWs.Cell(1, 1).Value, Is.EqualTo("Double referenced text"));
            });
        }

        [Test]
        public void Clear_dereferences_texts_in_the_range()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Double referenced text";
            ws.Cell("B2").Value = "Double referenced text";
            ws.Cell("C2").Value = "Single referenced text";

            Assert.That(sst.Count, Is.EqualTo(2));
            ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.Clear(new XLSheetRange(2, 2, 2, 3));
            Assert.Multiple(() =>
            {
                Assert.That(sst.Count, Is.EqualTo(1));
                Assert.That(ws.Cell("A1").Value, Is.EqualTo("Double referenced text"));
            });
        }

        [Test]
        public void DeleteAreaAndShiftLeft_dereferences_all_texts_deleted_area()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var ws = wb.AddWorksheet();
            ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
            ws.Cell("C1").Value = "Kept Single Reference"; // id 1
            ws.Cell("B1").Value = "Kept Double Reference"; // id 2
            ws.Cell("C3").Value = "Kept Double Reference"; // id 2

            ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(new XLSheetRange(2, 2, 3, 3));

            Assert.That(sst.Count, Is.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(sst[1], Is.EqualTo("Kept Single Reference"));
                Assert.That(sst[2], Is.EqualTo("Kept Double Reference"));
            });
        }

        [Test]
        public void DeleteAreaAndShiftUp_dereferences_all_texts_deleted_area()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var ws = wb.AddWorksheet();
            ws.Cell("B2").Value = "Deleted Single Reference"; // id 0
            ws.Cell("A3").Value = "Kept Single Reference"; // id 1
            ws.Cell("A2").Value = "Kept Double Reference"; // id 2
            ws.Cell("C3").Value = "Kept Double Reference"; // id 2

            ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.DeleteAreaAndShiftLeft(new XLSheetRange(2, 2, 3, 3));

            Assert.That(sst.Count, Is.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(sst[1], Is.EqualTo("Kept Single Reference"));
                Assert.That(sst[2], Is.EqualTo("Kept Double Reference"));
            });
        }

        [Test]
        public void InsertAreaAndShiftDown_dereferences_all_texts_in_pushed_out_range()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var ws = wb.AddWorksheet();

            ws.Cell("B2").Value = "Kept Single Reference"; // id 0
            ws.Cell("C1048576").Value = "Deleted Single Reference"; // id 1
            ws.Cell("B1048574").Value = "Kept Double Reference"; // id 2
            ws.Cell("B1048576").Value = "Kept Double Reference"; // id 2
            ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftDown(new XLSheetRange(3, 2, 4, 3));

            Assert.That(sst.Count, Is.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(sst[0], Is.EqualTo("Kept Single Reference"));
                Assert.That(sst[2], Is.EqualTo("Kept Double Reference"));
            });
        }

        [Test]
        public void InsertAreaAndShiftRight_dereferences_all_texts_in_pushed_out_range()
        {
            using var wb = new XLWorkbook();
            var sst = wb.SharedStringTable;
            var ws = wb.AddWorksheet();

            ws.Cell("B2").Value = "Kept Single Reference"; // id 0
            ws.Cell("XFD2").Value = "Deleted Single Reference"; // id 1
            ws.Cell("XFD3").Value = "Kept Double Reference"; // id 2
            ws.Cell("XFB3").Value = "Kept Double Reference"; // id 2
            ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.InsertAreaAndShiftRight(new XLSheetRange(2, 3, 3, 4));

            Assert.That(sst.Count, Is.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(sst[0], Is.EqualTo("Kept Single Reference"));
                Assert.That(sst[2], Is.EqualTo("Kept Double Reference"));
            });
        }
    }
}
