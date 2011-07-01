using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.PageSetup;
using ClosedXML_Examples.Ranges;
using ClosedXML_Examples.Rows;
using ClosedXML_Examples.Styles;

namespace ClosedXML_Examples
{
    public class CreateFiles
    {
        public static void CreateAllFiles()
        {
            new HelloWorld().Create(@"C:\Excel Files\Created\HelloWorld.xlsx");
            new BasicTable().Create(@"C:\Excel Files\Created\BasicTable.xlsx");

            new StyleExamples().Create();
            new ChangingBasicTable().Create(@"C:\Excel Files\Created\BasicTable_Modified.xlsx");
            new ShiftingRanges().Create(@"C:\Excel Files\Created\ShiftingRanges.xlsx");
            new ColumnSettings().Create(@"C:\Excel Files\Created\ColumnSettings.xlsx");
            new RowSettings().Create(@"C:\Excel Files\Created\RowSettings.xlsx");
            new MergeCells().Create(@"C:\Excel Files\Created\MergedCells.xlsx");
            new InsertRows().Create(@"C:\Excel Files\Created\InsertRows.xlsx");
            new InsertColumns().Create(@"C:\Excel Files\Created\InsertColumns.xlsx");
            new ColumnCollection().Create(@"C:\Excel Files\Created\ColumnCollection.xlsx");
            new DataTypes().Create(@"C:\Excel Files\Created\DataTypes.xlsx");
            new MultipleSheets().Create(@"C:\Excel Files\Created\MultipleSheets.xlsx");
            new RowCollection().Create(@"C:\Excel Files\Created\RowCollection.xlsx");
            new DefiningRanges().Create(@"C:\Excel Files\Created\DefiningRanges.xlsx");
            new ClearingRanges().Create(@"C:\Excel Files\Created\ClearingRanges.xlsx");
            new DeletingRanges().Create(@"C:\Excel Files\Created\DeletingRanges.xlsx");
            new Margins().Create(@"C:\Excel Files\Created\Margins.xlsx");
            new Page().Create(@"C:\Excel Files\Created\Page.xlsx");
            new HeaderFooters().Create(@"C:\Excel Files\Created\HeaderFooters.xlsx");
            new Sheets().Create(@"C:\Excel Files\Created\Sheets.xlsx");
            new SheetTab().Create(@"C:\Excel Files\Created\SheetTab.xlsx");
            new MultipleRanges().Create(@"C:\Excel Files\Created\MultipleRanges.xlsx");
            new StyleWorksheet().Create(@"C:\Excel Files\Created\StyleWorksheet.xlsx");
            new StyleRowsColumns().Create(@"C:\Excel Files\Created\StyleRowsColumns.xlsx");
            new InsertingDeletingRows().Create(@"C:\Excel Files\Created\InsertingDeletingRows.xlsx");
            new InsertingDeletingColumns().Create(@"C:\Excel Files\Created\InsertingDeletingColumns.xlsx");
            new DeletingColumns().Create(@"C:\Excel Files\Created\DeletingColumns.xlsx");
            new CellValues().Create(@"C:\Excel Files\Created\CellValues.xlsx");
            new LambdaExpressions().Create(@"C:\Excel Files\Created\LambdaExpressions.xlsx");
            new DefaultStyles().Create(@"C:\Excel Files\Created\DefaultStyles.xlsx");
            new TransposeRanges().Create(@"C:\Excel Files\Created\TransposeRanges.xlsx");
            new TransposeRangesPlus().Create(@"C:\Excel Files\Created\TransposeRangesPlus.xlsx");
            new MergeMoves().Create(@"C:\Excel Files\Created\MergedMoves.xlsx");
            new WorkbookProperties().Create(@"C:\Excel Files\Created\WorkbookProperties.xlsx");
            new AdjustToContents().Create(@"C:\Excel Files\Created\AdjustToContents.xlsx");
            new HideUnhide().Create(@"C:\Excel Files\Created\HideUnhide.xlsx");
            new Outline().Create(@"C:\Excel Files\Created\Outline.xlsx");
            new Formulas().Create(@"C:\Excel Files\Created\Formulas.xlsx");
            new Collections().Create(@"C:\Excel Files\Created\Collections.xlsx");
            new NamedRanges().Create(@"C:\Excel Files\Created\NamedRanges.xlsx");
            new CopyingRanges().Create(@"C:\Excel Files\Created\CopyingRanges.xlsx");
            new BlankCells().Create(@"C:\Excel Files\Created\BlankCells.xlsx");
            new TwoPages().Create(@"C:\Excel Files\Created\TwoPages.xlsx");
            new UsingColors().Create(@"C:\Excel Files\Created\UsingColors.xlsx");

            new ColumnCells().Create(@"C:\Excel Files\Created\ColumnCells.xlsx");
            new RowCells().Create(@"C:\Excel Files\Created\RowCells.xlsx");
            new FreezePanes().Create(@"C:\Excel Files\Created\FreezePanes.xlsx");
            new UsingTables().Create(@"C:\Excel Files\Created\UsingTables.xlsx");
            new ShowCase().Create(@"C:\Excel Files\Created\ShowCase.xlsx");
            new CopyingWorksheets().Create(@"C:\Excel Files\Created\CopyingWorksheets.xlsx");
            new InsertingTables().Create(@"C:\Excel Files\Created\InsertingTables.xlsx");
            new InsertingData().Create(@"C:\Excel Files\Created\InsertingData.xlsx");
            new Hyperlinks().Create(@"C:\Excel Files\Created\Hyperlinks.xlsx");
            new DataValidation().Create(@"C:\Excel Files\Created\DataValidation.xlsx");
            new HideSheets().Create(@"C:\Excel Files\Created\HideSheets.xlsx");
            new SheetProtection().Create(@"C:\Excel Files\Created\SheetProtection.xlsx");
            new AutoFilter().Create(@"C:\Excel Files\Created\AutoFilter.xlsx");
            new Sorting().Create(@"C:\Excel Files\Created\Sorting.xlsx");
            new SortExample().Create(@"C:\Excel Files\Created\SortExample.xlsx");
            new AddingDataSet().Create(@"C:\Excel Files\Created\AddingDataSet.xlsx");
            new AddingDataTableAsWorksheet().Create(@"C:\Excel Files\Created\AddingDataTableAsWorksheet.xlsx");
            new TabColors().Create(@"C:\Excel Files\Created\TabColors.xlsx");
            new ShiftingFormulas().Create(@"C:\Excel Files\Created\ShiftingFormulas.xlsx");
            new CopyingRowsAndColumns().Create(@"C:\Excel Files\Created\CopyingRowsAndColumns.xlsx");
            new UsingRichText().Create(@"C:\Excel Files\Created\UsingRichText.xlsx");
            new UsingPhonetics().Create(@"C:\Excel Files\Created\UsingPhonetics.xlsx");
        }
    }
}
