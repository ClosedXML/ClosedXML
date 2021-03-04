using ClosedXML.Examples.Columns;
using ClosedXML.Examples.Misc;
using ClosedXML.Examples.PageSetup;
using ClosedXML.Examples.Ranges;
using ClosedXML.Examples.Rows;
using ClosedXML.Examples.Styles;
using ClosedXML.Examples.Tables;
using System.IO;

namespace ClosedXML.Examples
{
    public class CreateFiles
    {
        public static void CreateAllFiles()
        {
            var path = Program.BaseCreatedDirectory;

            new HelloWorld().Create(Path.Combine(path, "HelloWorld.xlsx"));
            new BasicTable().Create(Path.Combine(path, "BasicTable.xlsx"));

            new StyleExamples().Create();
            new ChangingBasicTable().Create(Path.Combine(path, "BasicTable_Modified.xlsx"));
            new ShiftingRanges().Create(Path.Combine(path, "ShiftingRanges.xlsx"));
            new ColumnSettings().Create(Path.Combine(path, "ColumnSettings.xlsx"));
            new RowSettings().Create(Path.Combine(path, "RowSettings.xlsx"));
            new MergeCells().Create(Path.Combine(path, "MergedCells.xlsx"));
            new InsertRows().Create(Path.Combine(path, "InsertRows.xlsx"));
            new InsertColumns().Create(Path.Combine(path, "InsertColumns.xlsx"));
            new ColumnCollection().Create(Path.Combine(path, "ColumnCollection.xlsx"));
            new DataTypes().Create(Path.Combine(path, "DataTypes.xlsx"));
            new DataTypesUnderDifferentCulture().Create(Path.Combine(path, "DataTypesUnderDifferentCulture.xlsx"));
            new MultipleSheets().Create(Path.Combine(path, "MultipleSheets.xlsx"));
            new RowCollection().Create(Path.Combine(path, "RowCollection.xlsx"));
            new DefiningRanges().Create(Path.Combine(path, "DefiningRanges.xlsx"));
            new ClearingRanges().Create(Path.Combine(path, "ClearingRanges.xlsx"));
            new DeletingRanges().Create(Path.Combine(path, "DeletingRanges.xlsx"));
            new Margins().Create(Path.Combine(path, "Margins.xlsx"));
            new Page().Create(Path.Combine(path, "Page.xlsx"));
            new HeaderFooters().Create(Path.Combine(path, "HeaderFooters.xlsx"));
            new Sheets().Create(Path.Combine(path, "Sheets.xlsx"));
            new SheetTab().Create(Path.Combine(path, "SheetTab.xlsx"));
            new MultipleRanges().Create(Path.Combine(path, "MultipleRanges.xlsx"));
            new StyleWorksheet().Create(Path.Combine(path, "StyleWorksheet.xlsx"));
            new StyleRowsColumns().Create(Path.Combine(path, "StyleRowsColumns.xlsx"));
            new InsertingDeletingRows().Create(Path.Combine(path, "InsertingDeletingRows.xlsx"));
            new InsertingDeletingColumns().Create(Path.Combine(path, "InsertingDeletingColumns.xlsx"));
            new DeletingColumns().Create(Path.Combine(path, "DeletingColumns.xlsx"));
            new CellValues().Create(Path.Combine(path, "CellValues.xlsx"));
            new LambdaExpressions().Create(Path.Combine(path, "LambdaExpressions.xlsx"));
            new DefaultStyles().Create(Path.Combine(path, "DefaultStyles.xlsx"));
            new TransposeRanges().Create(Path.Combine(path, "TransposeRanges.xlsx"));
            new TransposeRangesPlus().Create(Path.Combine(path, "TransposeRangesPlus.xlsx"));
            new MergeMoves().Create(Path.Combine(path, "MergedMoves.xlsx"));
            new WorkbookProperties().Create(Path.Combine(path, "WorkbookProperties.xlsx"));
            new AdjustToContents().Create(Path.Combine(path, "AdjustToContents.xlsx"));
            new AdjustToContentsWithAutoFilter().Create(Path.Combine(path, "AdjustToContentsWithAutoFilter.xlsx"));
            new HideUnhide().Create(Path.Combine(path, "HideUnhide.xlsx"));
            new Outline().Create(Path.Combine(path, "Outline.xlsx"));
            new Formulas().Create(Path.Combine(path, "Formulas.xlsx"));
            new Collections().Create(Path.Combine(path, "Collections.xlsx"));
            new NamedRanges().Create(Path.Combine(path, "NamedRanges.xlsx"));
            new CopyingRanges().Create(Path.Combine(path, "CopyingRanges.xlsx"));
            new BlankCells().Create(Path.Combine(path, "BlankCells.xlsx"));
            new TwoPages().Create(Path.Combine(path, "TwoPages.xlsx"));
            new UsingColors().Create(Path.Combine(path, "UsingColors.xlsx"));

            new ColumnCells().Create(Path.Combine(path, "ColumnCells.xlsx"));
            new RowCells().Create(Path.Combine(path, "RowCells.xlsx"));
            new FreezePanes().Create(Path.Combine(path, "FreezePanes.xlsx"));
            new UsingTables().Create(Path.Combine(path, "UsingTables.xlsx"));
            new ResizingTables().Create(Path.Combine(path, "ResizingTables.xlsx"));
            new AddingRowToTables().Create(Path.Combine(path, "AddingRowToTables.xlsx"));
            new RightToLeft().Create(Path.Combine(path, "RightToLeft.xlsx"));
            new ShowCase().Create(Path.Combine(path, "ShowCase.xlsx"));
            new CopyingWorksheets().Create(Path.Combine(path, "CopyingWorksheets.xlsx"));
            new InsertingTables().Create(Path.Combine(path, "InsertingTables.xlsx"));
            new InsertingData().Create(Path.Combine(path, "InsertingData.xlsx"));
            new Hyperlinks().Create(Path.Combine(path, "Hyperlinks.xlsx"));
            new DataValidation().Create(Path.Combine(path, "DataValidation.xlsx"));
            new HideSheets().Create(Path.Combine(path, "HideSheets.xlsx"));
            new SheetProtection().Create(Path.Combine(path, "SheetProtection.xlsx"));
            new AutoFilter().Create(Path.Combine(path, "AutoFilter.xlsx"));
            new Sorting().Create(Path.Combine(path, "Sorting.xlsx"));
            new SortExample().Create(Path.Combine(path, "SortExample.xlsx"));
            new AddingDataSet().Create(Path.Combine(path, "AddingDataSet.xlsx"));
            new AddingDataTableAsWorksheet().Create(Path.Combine(path, "AddingDataTableAsWorksheet.xlsx"));
            new TabColors().Create(Path.Combine(path, "TabColors.xlsx"));
            new ShiftingFormulas().Create(Path.Combine(path, "ShiftingFormulas.xlsx"));
            new CopyingRowsAndColumns().Create(Path.Combine(path, "CopyingRowsAndColumns.xlsx"));
            new UsingRichText().Create(Path.Combine(path, "UsingRichText.xlsx"));
            new UsingPhonetics().Create(Path.Combine(path, "UsingPhonetics.xlsx"));
            new WalkingRanges().Create(Path.Combine(path, "CellMoves.xlsx"));
            new AddingComments().Create(Path.Combine(path, "AddingComments.xlsx"));
            new PivotTables().Create(Path.Combine(path, "PivotTables.xlsx"));
            new SheetViews().Create(Path.Combine(path, "SheetViews.xlsx"));
        }
    }
}
