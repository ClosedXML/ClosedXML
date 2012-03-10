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
            new HelloWorld().Create(@"D:\Excel Files\Created\HelloWorld.xlsx");
            new BasicTable().Create(@"D:\Excel Files\Created\BasicTable.xlsx");

            new StyleExamples().Create();
            new ChangingBasicTable().Create(@"D:\Excel Files\Created\BasicTable_Modified.xlsx");
            new ShiftingRanges().Create(@"D:\Excel Files\Created\ShiftingRanges.xlsx");
            new ColumnSettings().Create(@"D:\Excel Files\Created\ColumnSettings.xlsx");
            new RowSettings().Create(@"D:\Excel Files\Created\RowSettings.xlsx");
            new MergeCells().Create(@"D:\Excel Files\Created\MergedCells.xlsx");
            new InsertRows().Create(@"D:\Excel Files\Created\InsertRows.xlsx");
            new InsertColumns().Create(@"D:\Excel Files\Created\InsertColumns.xlsx");
            new ColumnCollection().Create(@"D:\Excel Files\Created\ColumnCollection.xlsx");
            new DataTypes().Create(@"D:\Excel Files\Created\DataTypes.xlsx");
            new MultipleSheets().Create(@"D:\Excel Files\Created\MultipleSheets.xlsx");
            new RowCollection().Create(@"D:\Excel Files\Created\RowCollection.xlsx");
            new DefiningRanges().Create(@"D:\Excel Files\Created\DefiningRanges.xlsx");
            new ClearingRanges().Create(@"D:\Excel Files\Created\ClearingRanges.xlsx");
            new DeletingRanges().Create(@"D:\Excel Files\Created\DeletingRanges.xlsx");
            new Margins().Create(@"D:\Excel Files\Created\Margins.xlsx");
            new Page().Create(@"D:\Excel Files\Created\Page.xlsx");
            new HeaderFooters().Create(@"D:\Excel Files\Created\HeaderFooters.xlsx");
            new Sheets().Create(@"D:\Excel Files\Created\Sheets.xlsx");
            new SheetTab().Create(@"D:\Excel Files\Created\SheetTab.xlsx");
            new MultipleRanges().Create(@"D:\Excel Files\Created\MultipleRanges.xlsx");
            new StyleWorksheet().Create(@"D:\Excel Files\Created\StyleWorksheet.xlsx");
            new StyleRowsColumns().Create(@"D:\Excel Files\Created\StyleRowsColumns.xlsx");
            new InsertingDeletingRows().Create(@"D:\Excel Files\Created\InsertingDeletingRows.xlsx");
            new InsertingDeletingColumns().Create(@"D:\Excel Files\Created\InsertingDeletingColumns.xlsx");
            new DeletingColumns().Create(@"D:\Excel Files\Created\DeletingColumns.xlsx");
            new CellValues().Create(@"D:\Excel Files\Created\CellValues.xlsx");
            new LambdaExpressions().Create(@"D:\Excel Files\Created\LambdaExpressions.xlsx");
            new DefaultStyles().Create(@"D:\Excel Files\Created\DefaultStyles.xlsx");
            new TransposeRanges().Create(@"D:\Excel Files\Created\TransposeRanges.xlsx");
            new TransposeRangesPlus().Create(@"D:\Excel Files\Created\TransposeRangesPlus.xlsx");
            new MergeMoves().Create(@"D:\Excel Files\Created\MergedMoves.xlsx");
            new WorkbookProperties().Create(@"D:\Excel Files\Created\WorkbookProperties.xlsx");
            new AdjustToContents().Create(@"D:\Excel Files\Created\AdjustToContents.xlsx");
            new HideUnhide().Create(@"D:\Excel Files\Created\HideUnhide.xlsx");
            new Outline().Create(@"D:\Excel Files\Created\Outline.xlsx");
            new Formulas().Create(@"D:\Excel Files\Created\Formulas.xlsx");
            new Collections().Create(@"D:\Excel Files\Created\Collections.xlsx");
            new NamedRanges().Create(@"D:\Excel Files\Created\NamedRanges.xlsx");
            new CopyingRanges().Create(@"D:\Excel Files\Created\CopyingRanges.xlsx");
            new BlankCells().Create(@"D:\Excel Files\Created\BlankCells.xlsx");
            new TwoPages().Create(@"D:\Excel Files\Created\TwoPages.xlsx");
            new UsingColors().Create(@"D:\Excel Files\Created\UsingColors.xlsx");

            new ColumnCells().Create(@"D:\Excel Files\Created\ColumnCells.xlsx");
            new RowCells().Create(@"D:\Excel Files\Created\RowCells.xlsx");
            new FreezePanes().Create(@"D:\Excel Files\Created\FreezePanes.xlsx");
            new UsingTables().Create(@"D:\Excel Files\Created\UsingTables.xlsx");
            new ShowCase().Create(@"D:\Excel Files\Created\ShowCase.xlsx");
            new CopyingWorksheets().Create(@"D:\Excel Files\Created\CopyingWorksheets.xlsx");
            new InsertingTables().Create(@"D:\Excel Files\Created\InsertingTables.xlsx");
            new InsertingData().Create(@"D:\Excel Files\Created\InsertingData.xlsx");
            new Hyperlinks().Create(@"D:\Excel Files\Created\Hyperlinks.xlsx");
            new DataValidation().Create(@"D:\Excel Files\Created\DataValidation.xlsx");
            new HideSheets().Create(@"D:\Excel Files\Created\HideSheets.xlsx");
            new SheetProtection().Create(@"D:\Excel Files\Created\SheetProtection.xlsx");
            new AutoFilter().Create(@"D:\Excel Files\Created\AutoFilter.xlsx");
            new Sorting().Create(@"D:\Excel Files\Created\Sorting.xlsx");
            new SortExample().Create(@"D:\Excel Files\Created\SortExample.xlsx");
            new AddingDataSet().Create(@"D:\Excel Files\Created\AddingDataSet.xlsx");
            new AddingDataTableAsWorksheet().Create(@"D:\Excel Files\Created\AddingDataTableAsWorksheet.xlsx");
            new TabColors().Create(@"D:\Excel Files\Created\TabColors.xlsx");
            new ShiftingFormulas().Create(@"D:\Excel Files\Created\ShiftingFormulas.xlsx");
            new CopyingRowsAndColumns().Create(@"D:\Excel Files\Created\CopyingRowsAndColumns.xlsx");
            new UsingRichText().Create(@"D:\Excel Files\Created\UsingRichText.xlsx");
            new UsingPhonetics().Create(@"D:\Excel Files\Created\UsingPhonetics.xlsx");
            new WalkingRanges().Create(@"D:\Excel Files\Created\CellMoves.xlsx");
            new AddingComments().Create(@"D:\Excel Files\Created\AddingComments.xlsx");
        }
    }
}
