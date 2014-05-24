using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.PageSetup;
using ClosedXML_Examples.Ranges;
using ClosedXML_Examples.Rows;
using ClosedXML_Examples.Styles;
using System.IO;

namespace ClosedXML_Examples
{
    public class CreateFiles
    {
        public static void CreateAllFiles()
        {
            var path = @"C:\ClosedXML_Tests\Created";
          
            new HelloWorld().Create(path + @"\HelloWorld.xlsx");
            new BasicTable().Create(path + @"\BasicTable.xlsx");

            new StyleExamples().Create();
            new ChangingBasicTable().Create(path + @"\BasicTable_Modified.xlsx");
            new ShiftingRanges().Create(path + @"\ShiftingRanges.xlsx");
            new ColumnSettings().Create(path + @"\ColumnSettings.xlsx");
            new RowSettings().Create(path + @"\RowSettings.xlsx");
            new MergeCells().Create(path + @"\MergedCells.xlsx");
            new InsertRows().Create(path + @"\InsertRows.xlsx");
            new InsertColumns().Create(path + @"\InsertColumns.xlsx");
            new ColumnCollection().Create(path + @"\ColumnCollection.xlsx");
            new DataTypes().Create(path + @"\DataTypes.xlsx");
            new MultipleSheets().Create(path + @"\MultipleSheets.xlsx");
            new RowCollection().Create(path + @"\RowCollection.xlsx");
            new DefiningRanges().Create(path + @"\DefiningRanges.xlsx");
            new ClearingRanges().Create(path + @"\ClearingRanges.xlsx");
            new DeletingRanges().Create(path + @"\DeletingRanges.xlsx");
            new Margins().Create(path + @"\Margins.xlsx");
            new Page().Create(path + @"\Page.xlsx");
            new HeaderFooters().Create(path + @"\HeaderFooters.xlsx");
            new Sheets().Create(path + @"\Sheets.xlsx");
            new SheetTab().Create(path + @"\SheetTab.xlsx");
            new MultipleRanges().Create(path + @"\MultipleRanges.xlsx");
            new StyleWorksheet().Create(path + @"\StyleWorksheet.xlsx");
            new StyleRowsColumns().Create(path + @"\StyleRowsColumns.xlsx");
            new InsertingDeletingRows().Create(path + @"\InsertingDeletingRows.xlsx");
            new InsertingDeletingColumns().Create(path + @"\InsertingDeletingColumns.xlsx");
            new DeletingColumns().Create(path + @"\DeletingColumns.xlsx");
            new CellValues().Create(path + @"\CellValues.xlsx");
            new LambdaExpressions().Create(path + @"\LambdaExpressions.xlsx");
            new DefaultStyles().Create(path + @"\DefaultStyles.xlsx");
            new TransposeRanges().Create(path + @"\TransposeRanges.xlsx");
            new TransposeRangesPlus().Create(path + @"\TransposeRangesPlus.xlsx");
            new MergeMoves().Create(path + @"\MergedMoves.xlsx");
            new WorkbookProperties().Create(path + @"\WorkbookProperties.xlsx");
            new AdjustToContents().Create(path + @"\AdjustToContents.xlsx");
            new HideUnhide().Create(path + @"\HideUnhide.xlsx");
            new Outline().Create(path + @"\Outline.xlsx");
            new Formulas().Create(path + @"\Formulas.xlsx");
            new Collections().Create(path + @"\Collections.xlsx");
            new NamedRanges().Create(path + @"\NamedRanges.xlsx");
            new CopyingRanges().Create(path + @"\CopyingRanges.xlsx");
            new BlankCells().Create(path + @"\BlankCells.xlsx");
            new TwoPages().Create(path + @"\TwoPages.xlsx");
            new UsingColors().Create(path + @"\UsingColors.xlsx");

            new ColumnCells().Create(path + @"\ColumnCells.xlsx");
            new RowCells().Create(path + @"\RowCells.xlsx");
            new FreezePanes().Create(path + @"\FreezePanes.xlsx");
            new UsingTables().Create(path + @"\UsingTables.xlsx");
            new ShowCase().Create(path + @"\ShowCase.xlsx");
            new CopyingWorksheets().Create(path + @"\CopyingWorksheets.xlsx");
            new InsertingTables().Create(path + @"\InsertingTables.xlsx");
            new InsertingData().Create(path + @"\InsertingData.xlsx");
            new Hyperlinks().Create(path + @"\Hyperlinks.xlsx");
            new DataValidation().Create(path + @"\DataValidation.xlsx");
            new HideSheets().Create(path + @"\HideSheets.xlsx");
            new SheetProtection().Create(path + @"\SheetProtection.xlsx");
            new AutoFilter().Create(path + @"\AutoFilter.xlsx");
            new Sorting().Create(path + @"\Sorting.xlsx");
            new SortExample().Create(path + @"\SortExample.xlsx");
            new AddingDataSet().Create(path + @"\AddingDataSet.xlsx");
            new AddingDataTableAsWorksheet().Create(path + @"\AddingDataTableAsWorksheet.xlsx");
            new TabColors().Create(path + @"\TabColors.xlsx");
            new ShiftingFormulas().Create(path + @"\ShiftingFormulas.xlsx");
            new CopyingRowsAndColumns().Create(path + @"\CopyingRowsAndColumns.xlsx");
            new UsingRichText().Create(path + @"\UsingRichText.xlsx");
            new UsingPhonetics().Create(path + @"\UsingPhonetics.xlsx");
            new WalkingRanges().Create(path + @"\CellMoves.xlsx");
            new AddingComments().Create(path + @"\AddingComments.xlsx");
        }
    }
}
