using ClosedXML_Examples;
using ClosedXML_Examples.Misc;
using NUnit.Framework;

namespace ClosedXML_Tests.Examples
{
    [TestFixture]
    public class MiscTests
    {
        [Test]
        public void AddingDataSet()
        {
            TestHelper.RunTestExample<AddingDataSet>(@"Misc\AddingDataSet.xlsx");
        }

        [Test]
        public void AddingDataTableAsWorksheet()
        {
            TestHelper.RunTestExample<AddingDataTableAsWorksheet>(@"Misc\AddingDataTableAsWorksheet.xlsx");
        }

        [Test]
        public void AdjustToContents()
        {
            TestHelper.RunTestExample<AdjustToContents>(@"Misc\AdjustToContents.xlsx");
        }

        [Test]
        public void AdjustToContentsWithAutoFilter()
        {
            TestHelper.RunTestExample<AdjustToContentsWithAutoFilter>(@"Misc\AdjustToContentsWithAutoFilter.xlsx");
        }

        [Test]
        public void AutoFilter()
        {
            TestHelper.RunTestExample<AutoFilter>(@"Misc\AutoFilter.xlsx");
        }

        [Test]
        public void BasicTable()
        {
            TestHelper.RunTestExample<BasicTable>(@"Misc\BasicTable.xlsx");
        }

        [Test]
        public void BlankCells()
        {
            TestHelper.RunTestExample<BlankCells>(@"Misc\BlankCells.xlsx");
        }

        [Test]
        public void CellValues()
        {
            TestHelper.RunTestExample<CellValues>(@"Misc\CellValues.xlsx");
        }

        [Test]
        public void Collections()
        {
            TestHelper.RunTestExample<Collections>(@"Misc\Collections.xlsx");
        }

        [Test]
        public void CopyingRowsAndColumns()
        {
            TestHelper.RunTestExample<CopyingRowsAndColumns>(@"Misc\CopyingRowsAndColumns.xlsx");
        }

        [Test]
        public void CopyingWorksheets()
        {
            TestHelper.RunTestExample<CopyingWorksheets>(@"Misc\CopyingWorksheets.xlsx");
        }

        [Test]
        public void DataTypes()
        {
            TestHelper.RunTestExample<DataTypes>(@"Misc\DataTypes.xlsx");
        }

        [Test]
        public void DataTypesUnderDifferentCulture()
        {
            TestHelper.RunTestExample<DataTypesUnderDifferentCulture>(@"Misc\DataTypesUnderDifferentCulture.xlsx");
        }

        [Test]
        public void DataValidation()
        {
            TestHelper.RunTestExample<DataValidation>(@"Misc\DataValidation.xlsx");
        }

        [Test]
        public void Formulas()
        {
            TestHelper.RunTestExample<Formulas>(@"Misc\Formulas.xlsx");
        }

        [Test]
        public void FormulasWithEvaluation()
        {
            TestHelper.RunTestExample<FormulasWithEvaluation>(@"Misc\FormulasWithEvaluation.xlsx", true);
        }

        [Test]
        public void FreezePanes()
        {
            TestHelper.RunTestExample<FreezePanes>(@"Misc\FreezePanes.xlsx");
        }

        [Test]
        public void HideSheets()
        {
            TestHelper.RunTestExample<HideSheets>(@"Misc\HideSheets.xlsx");
        }

        [Test]
        public void HideUnhide()
        {
            TestHelper.RunTestExample<HideUnhide>(@"Misc\HideUnhide.xlsx");
        }

        [Test]
        public void Hyperlinks()
        {
            TestHelper.RunTestExample<Hyperlinks>(@"Misc\Hyperlinks.xlsx");
        }

        [Test]
        public void InsertingData()
        {
            TestHelper.RunTestExample<InsertingData>(@"Misc\InsertingData.xlsx");
        }


        [Test]
        public void LambdaExpressions()
        {
            TestHelper.RunTestExample<LambdaExpressions>(@"Misc\LambdaExpressions.xlsx");
        }

        [Test]
        public void MergeCells()
        {
            TestHelper.RunTestExample<MergeCells>(@"Misc\MergeCells.xlsx");
        }

        [Test]
        public void MergeMoves()
        {
            TestHelper.RunTestExample<MergeMoves>(@"Misc\MergeMoves.xlsx");
        }

        [Test]
        public void Outline()
        {
            TestHelper.RunTestExample<Outline>(@"Misc\Outline.xlsx");
        }

        [Test]
        public void RightToLeft()
        {
            TestHelper.RunTestExample<RightToLeft>(@"Misc\RightToLeft.xlsx");
        }

        [Test]
        public void SheetProtection()
        {
            TestHelper.RunTestExample<SheetProtection>(@"Misc\SheetProtection.xlsx");
        }

        [Test]
        public void SheetViews()
        {
            TestHelper.RunTestExample<SheetViews>(@"Misc\SheetViews.xlsx");
        }

        [Test]
        public void ShiftingFormulas()
        {
            TestHelper.RunTestExample<ShiftingFormulas>(@"Misc\ShiftingFormulas.xlsx");
        }

        [Test]
        public void ShowCase()
        {
            TestHelper.RunTestExample<ShowCase>(@"Misc\ShowCase.xlsx");
        }

        [Test]
        public void TabColors()
        {
            TestHelper.RunTestExample<TabColors>(@"Misc\TabColors.xlsx");
        }

        [Test]
        public void WorkbookProperties()
        {
            TestHelper.RunTestExample<WorkbookProperties>(@"Misc\WorkbookProperties.xlsx");
        }

        [Test]
        public void WorkbookProtection()
        {
            TestHelper.RunTestExample<WorkbookProtection>(@"Misc\WorkbookProtection.xlsx");
        }
    }
}
