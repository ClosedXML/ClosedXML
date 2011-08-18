using ClosedXML_Examples;
using ClosedXML_Examples.Misc;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class MiscTests
    {
        [TestMethod]
        public void AddingDataSet()
        {
            TestHelper.RunTestExample<AddingDataSet>(@"Misc\AddingDataSet.xlsx");
        }
        [TestMethod]
        public void AddingDataTableAsWorksheet()
        {
            TestHelper.RunTestExample<AddingDataTableAsWorksheet>(@"Misc\AddingDataTableAsWorksheet.xlsx");
        }
        [TestMethod]
        public void AdjustToContents()
        {
            TestHelper.RunTestExample<AdjustToContents>(@"Misc\AdjustToContents.xlsx");
        }
        [TestMethod]
        public void AutoFilter()
        {
            TestHelper.RunTestExample<AutoFilter>(@"Misc\AutoFilter.xlsx");
        }
        [TestMethod]
        public void BlankCells()
        {
            TestHelper.RunTestExample<BlankCells>(@"Misc\BlankCells.xlsx");
        }
        [TestMethod]
        public void CellValues()
        {
            TestHelper.RunTestExample<CellValues>(@"Misc\CellValues.xlsx");
        }
        [TestMethod]
        public void Collections()
        {
            TestHelper.RunTestExample<Collections>(@"Misc\Collections.xlsx");
        }
        [TestMethod]
        public void CopyingWorksheets()
        {
            TestHelper.RunTestExample<CopyingWorksheets>(@"Misc\CopyingWorksheets.xlsx");
        }
        [TestMethod]
        public void DataTypes()
        {
            TestHelper.RunTestExample<DataTypes>(@"Misc\DataTypes.xlsx");
        }
        [TestMethod]
        public void DataValidation()
        {
            TestHelper.RunTestExample<DataValidation>(@"Misc\DataValidation.xlsx");
        }
        [TestMethod]
        public void Formulas()
        {
            TestHelper.RunTestExample<Formulas>(@"Misc\Formulas.xlsx");
        }
        [TestMethod]
        public void FreezePanes()
        {
            TestHelper.RunTestExample<FreezePanes>(@"Misc\FreezePanes.xlsx");
        }
        [TestMethod]
        public void HideSheets()
        {
            TestHelper.RunTestExample<HideSheets>(@"Misc\HideSheets.xlsx");
        }
        [TestMethod]
        public void HideUnhide()
        {
            TestHelper.RunTestExample<HideUnhide>(@"Misc\HideUnhide.xlsx");
        }
        [TestMethod]
        public void Hyperlinks()
        {
            TestHelper.RunTestExample<Hyperlinks>(@"Misc\Hyperlinks.xlsx");
        }
        [TestMethod]
        public void InsertingData()
        {
            TestHelper.RunTestExample<InsertingData>(@"Misc\InsertingData.xlsx");
        }
        [TestMethod]
        public void InsertingTables()
        {
            TestHelper.RunTestExample<InsertingTables>(@"Misc\InsertingTables.xlsx");
        }
        [TestMethod]
        public void LambdaExpressions()
        {
            TestHelper.RunTestExample<LambdaExpressions>(@"Misc\LambdaExpressions.xlsx");
        }
        [TestMethod]
        public void MergeCells()
        {
            TestHelper.RunTestExample<MergeCells>(@"Misc\MergeCells.xlsx");
        }
        [TestMethod]
        public void MergeMoves()
        {
            TestHelper.RunTestExample<MergeMoves>(@"Misc\MergeMoves.xlsx");
        }
        [TestMethod]
        public void MultipleSheets()
        {
            TestHelper.RunTestExample<MultipleSheets>(@"Misc\MultipleSheets.xlsx");
        }
        [TestMethod]
        public void Outline()
        {
            TestHelper.RunTestExample<Outline>(@"Misc\Outline.xlsx");
        }
        [TestMethod]
        public void SheetProtection()
        {
            TestHelper.RunTestExample<SheetProtection>(@"Misc\SheetProtection.xlsx");
        }
        [TestMethod]
        public void ShiftingFormulas()
        {
            TestHelper.RunTestExample<ShiftingFormulas>(@"Misc\ShiftingFormulas.xlsx");
        }
        [TestMethod]
        public void ShowCase()
        {
            TestHelper.RunTestExample<ShowCase>(@"Misc\ShowCase.xlsx");
        }
        [TestMethod]
        public void TabColors()
        {
            TestHelper.RunTestExample<TabColors>(@"Misc\TabColors.xlsx");
        }
        [TestMethod]
        public void WorkbookProperties()
        {
            TestHelper.RunTestExample<WorkbookProperties>(@"Misc\WorkbookProperties.xlsx");
        }
        [TestMethod]
        public void CopyingRowsAndColumns()
        {
            TestHelper.RunTestExample<CopyingRowsAndColumns>(@"Misc\CopyingRowsAndColumns.xlsx");
        }
        [TestMethod]
        public void BasicTable()
        {
            TestHelper.RunTestExample<BasicTable>(@"Misc\BasicTable.xlsx");
        }
    }
}