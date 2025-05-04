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
            Assert.That(slice[point], Is.EqualTo(1));
        }

        [Test]
        public void Setting_Value_To_Default_Clears_Element()
        {
            var slice = new Slice<int>();
            var point = new XLSheetPoint(574, 241);
            slice.Set(point, 1);
            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(574));
                Assert.That(slice.MaxColumn, Is.EqualTo(241));
            });

            slice.Set(point, 0);

            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(0));
                Assert.That(slice.MaxColumn, Is.EqualTo(0));
            });
        }

        [Test]
        public void Keeps_Track_Of_Max_Used_Coordinates()
        {
            var slice = new Slice<int>();
            slice.Set(54, 32, 1);
            slice.Set(140, 32, 1);
            slice.Set(140, 72, 1);

            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(140));
                Assert.That(slice.MaxColumn, Is.EqualTo(72));
            });

            slice.Set(140, 72, 0);

            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(140));
                Assert.That(slice.MaxColumn, Is.EqualTo(32));
            });

            slice.Set(140, 32, 0);

            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(54));
                Assert.That(slice.MaxColumn, Is.EqualTo(32));
            });

            slice.Set(54, 32, 0);

            Assert.Multiple(() =>
            {
                Assert.That(slice.MaxRow, Is.EqualTo(0));
                Assert.That(slice.MaxColumn, Is.EqualTo(0));
            });
        }

        [Test]
        public void Keeps_Track_Of_Used_Rows()
        {
            var slice = new Slice<int>();
            Assert.IsEmpty(slice.UsedRows);

            slice.Set(new XLSheetPoint(1, 1), 1);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1 }));

            slice.Set(new XLSheetPoint(70, 1), 1);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1, 70 }));

            slice.Set(new XLSheetPoint(35, 1), 1);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1, 35, 70 }));

            slice.Set(new XLSheetPoint(35, 2), 1);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1, 35, 70 }));

            slice.Set(new XLSheetPoint(35, 1), 0);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1, 35, 70 }));

            slice.Set(new XLSheetPoint(35, 2), 0);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 1, 70 }));

            slice.Set(new XLSheetPoint(1, 1), 0);
            Assert.That(slice.UsedRows, Is.EquivalentTo(new[] { 70 }));

            slice.Set(new XLSheetPoint(70, 1), 0);
            Assert.IsEmpty(slice.UsedRows);
        }

        [Test]
        public void Keeps_Track_Of_Used_Columns()
        {
            var slice = new Slice<int>();
            Assert.IsEmpty(slice.UsedColumns);

            slice.Set(new XLSheetPoint(1, 5), 1);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 5 }));

            slice.Set(new XLSheetPoint(1, 750), 1);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 5, 750 }));

            slice.Set(new XLSheetPoint(1, 90), 1);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 5, 90, 750 }));

            slice.Set(new XLSheetPoint(2, 5), 1);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 5, 90, 750 }));

            slice.Set(new XLSheetPoint(1, 5), 0);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 5, 90, 750 }));

            slice.Set(new XLSheetPoint(2, 5), 0);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 90, 750 }));

            slice.Set(new XLSheetPoint(1, 750), 0);
            Assert.That(slice.UsedColumns, Is.EquivalentTo(new[] { 90 }));

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
            Assert.Multiple(() =>
            {
                Assert.That(slice[outsideAddress], Is.EqualTo(1));
                Assert.That(slice[firstCorner], Is.EqualTo(0));
                Assert.That(slice[insideAddress], Is.EqualTo(0));
                Assert.That(slice[lastCorner], Is.EqualTo(0));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(slice[3, 1], Is.EqualTo(1));
                Assert.That(slice[5, 1], Is.EqualTo(2));
                Assert.That(slice[XLHelper.MaxRowNumber, 2], Is.EqualTo(0));
                Assert.That(slice[outsideAddress], Is.EqualTo(4));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(slice[1, 3], Is.EqualTo(1));
                Assert.That(slice[1, 5], Is.EqualTo(2));
                Assert.That(slice[purgedAddress], Is.EqualTo(0));
                Assert.That(slice[outsideAddress], Is.EqualTo(4));
            });
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
            Assert.Multiple(() =>
            {
                Assert.That(slice[firstCorner], Is.EqualTo(0));
                Assert.That(slice[secondCorner], Is.EqualTo(0));
                Assert.That(slice[belowAddress.Row - deleteArea.Height, belowAddress.Column], Is.EqualTo(5));
                Assert.That(slice[aboveAddress], Is.EqualTo(1));
                Assert.That(slice[rightAddress], Is.EqualTo(4));
                Assert.That(slice[leftAddress], Is.EqualTo(6));
            });
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
            Assert.Multiple(() =>
            {
                Assert.That(slice[firstCorner], Is.EqualTo(0));
                Assert.That(slice[secondCorner], Is.EqualTo(0));
                Assert.That(slice[rightAddress.Row, rightAddress.Column - deleteArea.Width], Is.EqualTo(5));
                Assert.That(slice[leftAddress], Is.EqualTo(1));
                Assert.That(slice[belowAddress], Is.EqualTo(4));
                Assert.That(slice[aboveAddress], Is.EqualTo(6));
            });
        }
    }
}
