using ClosedXML.Examples;
using ClosedXML.Examples.Misc;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class MiscTests
    {
        [Test]
        public void AddingDataSet()
        {
            TestHelper.RunTestExample<AddingDataSet>(@"Misc\AddingDataSet.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void AddingDataTableAsWorksheet()
        {
            TestHelper.RunTestExample<AddingDataTableAsWorksheet>(@"Misc\AddingDataTableAsWorksheet.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void AdjustToContents()
        {
            // adjusted cell height depends on fonts available on the test system
            var allowedDiff = "/xl/worksheets/sheet1.xml :NonEqual\n/xl/worksheets/sheet4.xml :NonEqual\n/xl/worksheets/sheet5.xml :NonEqual\n";

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                allowedDiff = null;
            }

            TestHelper.RunTestExample<AdjustToContents>(@"Misc\AdjustToContents.xlsx", false, allowedDiff, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void AdjustToContentsWithAutoFilter()
        {
            TestHelper.RunTestExample<AdjustToContentsWithAutoFilter>(@"Misc\AdjustToContentsWithAutoFilter.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void AutoFilter()
        {
            TestHelper.RunTestExample<AutoFilter>(@"Misc\AutoFilter.xlsx");
        }

        [Test]
        public void BasicTable()
        {
            TestHelper.RunTestExample<BasicTable>(@"Misc\BasicTable.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void BlankCells()
        {
            TestHelper.RunTestExample<BlankCells>(@"Misc\BlankCells.xlsx");
        }

        [Test]
        public void CellValues()
        {
            TestHelper.RunTestExample<CellValues>(@"Misc\CellValues.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void Collections()
        {
            TestHelper.RunTestExample<Collections>(@"Misc\Collections.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void CopyingRowsAndColumns()
        {
            TestHelper.RunTestExample<CopyingRowsAndColumns>(@"Misc\CopyingRowsAndColumns.xlsx");
        }

        [Test]
        public void CopyingWorksheets()
        {
            TestHelper.RunTestExample<CopyingWorksheets>(@"Misc\CopyingWorksheets.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void DataTypes()
        {
            TestHelper.RunTestExample<DataTypes>(@"Misc\DataTypes.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void DataTypesUnderDifferentCulture()
        {
            TestHelper.RunTestExample<DataTypesUnderDifferentCulture>(@"Misc\DataTypesUnderDifferentCulture.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void DataValidation()
        {
            TestHelper.RunTestExample<DataValidation>(@"Misc\DataValidation.xlsx");
        }

        [Test]
        public void DataValidationDecimal()
        {
            TestHelper.RunTestExample<DataValidationDecimal>(@"Misc\DataValidationDecimal.xlsx");
        }

        [Test]
        public void DataValidationWholeNumber()
        {
            TestHelper.RunTestExample<DataValidationWholeNumber>(@"Misc\DataValidationWholeNumber.xlsx");
        }

        [Test]
        public void DataValidationTextLength()
        {
            TestHelper.RunTestExample<DataValidationTextLength>(@"Misc\DataValidationTextLength.xlsx");
        }

        [Test]
        public void DataValidationDate()
        {
            TestHelper.RunTestExample<DataValidationDate>(@"Misc\DataValidationDate.xlsx");
        }

        [Test]
        public void DataValidationTime()
        {
            TestHelper.RunTestExample<DataValidationTime>(@"Misc\DataValidationTime.xlsx");
        }

        [Test]
        public void Formulas()
        {
            TestHelper.RunTestExample<Formulas>(@"Misc\Formulas.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void FormulasWithEvaluation()
        {
            TestHelper.RunTestExample<FormulasWithEvaluation>(@"Misc\FormulasWithEvaluation.xlsx", true, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
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
            TestHelper.RunTestExample<Hyperlinks>(@"Misc\Hyperlinks.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void InsertingData()
        {
            var exepectation = @"Misc\InsertingData.xlsx";

#if NETFRAMEWORK
             exepectation = @"Misc\InsertingDataNetFramework.xlsx";
#endif

            TestHelper.RunTestExample<InsertingData>(exepectation, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void LambdaExpressions()
        {
            TestHelper.RunTestExample<LambdaExpressions>(@"Misc\LambdaExpressions.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
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
            TestHelper.RunTestExample<SheetProtection>(@"Misc\SheetProtection.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
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
            TestHelper.RunTestExample<ShowCase>(@"Misc\ShowCase.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
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