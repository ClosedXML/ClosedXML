using ClosedXML_Examples;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.Ranges;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class RangesTests
    {
        [TestMethod]
        public void ClearingRanges()
        {
            TestHelper.RunTestExample<ClearingRanges>(@"Ranges\ClearingRanges.xlsx");
        }
        [TestMethod]
        public void CopyingRanges()
        {
            TestHelper.RunTestExample<CopyingRanges>(@"Ranges\CopyingRanges.xlsx");
        }
        [TestMethod]
        public void DefiningRanges()
        {
            TestHelper.RunTestExample<DefiningRanges>(@"Ranges\DefiningRanges.xlsx");
        }
        [TestMethod]
        public void DeletingRanges()
        {
            TestHelper.RunTestExample<DeletingRanges>(@"Ranges\DeletingRanges.xlsx");
        }
        [TestMethod]
        public void InsertingDeletingColumns()
        {
            TestHelper.RunTestExample<InsertingDeletingColumns>(@"Ranges\InsertingDeletingColumns.xlsx");
        }
        [TestMethod]
        public void InsertingDeletingRows()
        {
            TestHelper.RunTestExample<InsertingDeletingRows>(@"Ranges\InsertingDeletingRows.xlsx");
        }
        [TestMethod]
        public void MultipleRanges()
        {
            TestHelper.RunTestExample<MultipleRanges>(@"Ranges\MultipleRanges.xlsx");
        }
        [TestMethod]
        public void NamedRanges()
        {
            TestHelper.RunTestExample<NamedRanges>(@"Ranges\NamedRanges.xlsx");
        }
        [TestMethod]
        public void ShiftingRanges()
        {
            TestHelper.RunTestExample<ShiftingRanges>(@"Ranges\ShiftingRanges.xlsx");
        }
        [TestMethod]
        public void SortExample()
        {
            TestHelper.RunTestExample<SortExample>(@"Ranges\SortExample.xlsx");
        }
        [TestMethod]
        public void Sorting()
        {
            TestHelper.RunTestExample<Sorting>(@"Ranges\Sorting.xlsx");
        }
        [TestMethod]
        public void TransposeRanges()
        {
            TestHelper.RunTestExample<TransposeRanges>(@"Ranges\TransposeRanges.xlsx");
        }
        [TestMethod]
        public void TransposeRangesPlus()
        {
            TestHelper.RunTestExample<TransposeRangesPlus>(@"Ranges\TransposeRangesPlus.xlsx");
        }
        [TestMethod]
        public void UsingTables()
        {
            TestHelper.RunTestExample<UsingTables>(@"Ranges\UsingTables.xlsx");
        }

    }
}