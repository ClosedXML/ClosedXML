using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class SliceTests
    {
        [Test]
        public void Stores_Values()
        {
            var slice = new Slice<int>();
            var point = new XLSheetPoint(574, 241);
            slice.Set(point, 1);
            Assert.AreEqual(1, slice[point]);
        }

        [Test]
        public void Setting_Value_To_Default_Clears_Element()
        {
            var slice = new Slice<int>();
            var point = new XLSheetPoint(574, 241);
            slice.Set(point, 1);
            Assert.AreEqual(574, slice.MaxRow);
            Assert.AreEqual(241, slice.MaxColumn);

            slice.Set(point, 0);

            Assert.AreEqual(0, slice.MaxRow);
            Assert.AreEqual(0, slice.MaxColumn);
        }

        [Test]
        public void Keeps_Track_Of_Max_Used_Coordinates()
        {
            var slice = new Slice<int>();
            slice.Set(54, 32, 1);
            slice.Set(140, 32, 1);
            slice.Set(140, 72, 1);

            Assert.AreEqual(140, slice.MaxRow);
            Assert.AreEqual(72, slice.MaxColumn);

            slice.Set(140, 72, 0);

            Assert.AreEqual(140, slice.MaxRow);
            Assert.AreEqual(32, slice.MaxColumn);

            slice.Set(140, 32, 0);

            Assert.AreEqual(54, slice.MaxRow);
            Assert.AreEqual(32, slice.MaxColumn);

            slice.Set(54, 32, 0);

            Assert.AreEqual(0, slice.MaxRow);
            Assert.AreEqual(0, slice.MaxColumn);
        }

        [Test]
        public void Keeps_Track_Of_Used_Rows()
        {
            var slice = new Slice<int>();
            Assert.IsEmpty(slice.UsedRows);

            slice.Set(new XLSheetPoint(1, 1), 1);
            CollectionAssert.AreEquivalent(new[] { 1 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(70, 1), 1);
            CollectionAssert.AreEquivalent(new[] { 1, 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(35, 1), 1);
            CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(35, 2), 1);
            CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(35, 1), 0);
            CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(35, 2), 0);
            CollectionAssert.AreEquivalent(new[] { 1, 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(1, 1), 0);
            CollectionAssert.AreEquivalent(new[] { 70 }, slice.UsedRows);

            slice.Set(new XLSheetPoint(70, 1), 0);
            Assert.IsEmpty(slice.UsedRows);
        }

        [Test]
        public void Keeps_Track_Of_Used_Columns()
        {
            var slice = new Slice<int>();
            Assert.IsEmpty(slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 5), 1);
            CollectionAssert.AreEquivalent(new[] { 5 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 750), 1);
            CollectionAssert.AreEquivalent(new[] { 5, 750 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 90), 1);
            CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(2, 5), 1);
            CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 5), 0);
            CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(2, 5), 0);
            CollectionAssert.AreEquivalent(new[] { 90, 750 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 750), 0);
            CollectionAssert.AreEquivalent(new[] { 90 }, slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 90), 0);
            Assert.IsEmpty(slice.UsedColumns);
        }

        [Test]
        public void Clear_Range_Sets_Values_To_Default()
        {
            var slice = new Slice<int>();
            var outsideAddress = new XLSheetPoint(1, 1);
            slice.Set(outsideAddress, 1);
            var firstCorner = new XLSheetPoint(50, 20);
            slice.Set(firstCorner, 1);
            var insideAddress = new XLSheetPoint(55, 22);
            slice.Set(insideAddress, 1);
            var lastCorner = new XLSheetPoint(60, 30);
            slice.Set(lastCorner, 1);

            slice.Clear(new XLSheetRange(firstCorner, lastCorner));
            Assert.AreEqual(1, slice[outsideAddress]);
            Assert.AreEqual(0, slice[firstCorner]);
            Assert.AreEqual(0, slice[insideAddress]);
            Assert.AreEqual(0, slice[lastCorner]);
        }

        [Test]
        public void InsertAreaAndShiftDown_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
        {
            var slice = new Slice<int>();
            slice.Set(1, 1, 1);
            slice.Set(3, 1, 2);
            var purgedAddress = new XLSheetPoint(XLHelper.MaxRowNumber, 2);
            slice.Set(purgedAddress, 3);

            var outsideAddress = new XLSheetPoint(1, 3);
            slice.Set(outsideAddress, 4);

            slice.InsertAreaAndShiftDown(new XLSheetRange(new XLSheetPoint(1, 1), new XLSheetPoint(2, 2)));

            Assert.AreEqual(1, slice[3, 1]);
            Assert.AreEqual(2, slice[5, 1]);
            Assert.AreEqual(0, slice[XLHelper.MaxRowNumber, 2]);
            Assert.AreEqual(4, slice[outsideAddress]);
        }

        [Test]
        public void InsertAreaAndShiftRight_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
        {
            var slice = new Slice<int>();
            slice.Set(1, 1, 1);
            slice.Set(1, 3, 2);
            var purgedAddress = new XLSheetPoint(2, XLHelper.MaxColumnNumber);
            slice.Set(purgedAddress, 3);

            var outsideAddress = new XLSheetPoint(3, 1);
            slice.Set(outsideAddress, 4);

            slice.InsertAreaAndShiftRight(new XLSheetRange(new XLSheetPoint(1, 1), new XLSheetPoint(2, 2)));

            Assert.AreEqual(1, slice[1, 3]);
            Assert.AreEqual(2, slice[1, 5]);
            Assert.AreEqual(0, slice[purgedAddress]);
            Assert.AreEqual(4, slice[outsideAddress]);
        }

        [Test]
        public void DeleteAreaAndShiftUp_Moves_Area_Cells_Up()
        {
            var slice = new Slice<int>();
            var aboveAddress = new XLSheetPoint(1, 3);
            slice.Set(aboveAddress, 1);
            var firstCorner = new XLSheetPoint(2, 2);
            slice.Set(firstCorner, 2);
            var secondCorner = new XLSheetPoint(4, 5);
            slice.Set(secondCorner, 3);
            var rightAddress = new XLSheetPoint(3, 6);
            slice.Set(rightAddress, 4);
            var belowAddress = new XLSheetPoint(5, 3);
            slice.Set(belowAddress, 5);
            var leftAddress = new XLSheetPoint(3, 1);
            slice.Set(leftAddress, 6);

            var deleteArea = new XLSheetRange(firstCorner, secondCorner);
            slice.DeleteAreaAndShiftUp(deleteArea);
            Assert.AreEqual(0, slice[firstCorner]);
            Assert.AreEqual(0, slice[secondCorner]);
            Assert.AreEqual(5, slice[belowAddress.Row - deleteArea.Height, belowAddress.Column]);
            Assert.AreEqual(1, slice[aboveAddress]);
            Assert.AreEqual(4, slice[rightAddress]);
            Assert.AreEqual(6, slice[leftAddress]);
        }

        [Test]
        public void DeleteAreaAndShiftLeft_Moves_Area_Cells_Left()
        {
            var slice = new Slice<int>();
            var leftAddress = new XLSheetPoint(3, 1);
            slice.Set(leftAddress, 1);
            var firstCorner = new XLSheetPoint(2, 2);
            slice.Set(firstCorner, 2);
            var secondCorner = new XLSheetPoint(5, 4);
            slice.Set(secondCorner, 3);
            var belowAddress = new XLSheetPoint(6, 3);
            slice.Set(belowAddress, 4);
            var rightAddress = new XLSheetPoint(3, 5);
            slice.Set(rightAddress, 5);
            var aboveAddress = new XLSheetPoint(1, 3);
            slice.Set(aboveAddress, 6);

            var deleteArea = new XLSheetRange(firstCorner, secondCorner);
            slice.DeleteAreaAndShiftLeft(deleteArea);
            Assert.AreEqual(0, slice[firstCorner]);
            Assert.AreEqual(0, slice[secondCorner]);
            Assert.AreEqual(5, slice[rightAddress.Row, rightAddress.Column - deleteArea.Width]);
            Assert.AreEqual(1, slice[leftAddress]);
            Assert.AreEqual(4, slice[belowAddress]);
            Assert.AreEqual(6, slice[aboveAddress]);
        }
    }
}
