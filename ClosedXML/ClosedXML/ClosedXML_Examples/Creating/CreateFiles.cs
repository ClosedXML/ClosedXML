using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML_Examples.Styles;
using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Rows;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.Ranges;
using ClosedXML_Examples.PageSetup;

namespace ClosedXML_Examples
{
    public class CreateFiles
    {
        public static void CreateAllFiles()
        {
            new HelloWorld().Create(@"C:\Excel Files\Created\HelloWorld.xlsx");
            new BasicTable().Create(@"C:\Excel Files\Created\BasicTable.xlsx");
            new StyleExamples().Create();
            new ChangingBasicTable().Create();
            new ShiftingRanges().Create();
            new ColumnSettings().Create(@"C:\Excel Files\Created\ColumnSettings.xlsx");
            new RowSettings().Create(@"C:\Excel Files\Created\RowSettings.xlsx");
            new MergeCells().Create(@"C:\Excel Files\Created\MergedCells.xlsx");
            new InsertRows().Create(@"C:\Excel Files\Created\InsertRows.xlsx");
            new InsertColumns().Create(@"C:\Excel Files\Created\InsertColumns.xlsx");
            new ColumnCollection().Create(@"C:\Excel Files\Created\ColumnCollection.xlsx");
            new DataTypes().Create(@"C:\Excel Files\Created\DataTypes.xlsx");
            new MultipleSheets().Create();
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
            new LambdaExpressions().Create();
            new DefaultStyles().Create(@"C:\Excel Files\Created\DefaultStyles.xlsx");
            new TransposeRanges().Create();
            new TransposeRangesPlus().Create();
            new MergeMoves().Create();
            new WorkbookProperties().Create(@"C:\Excel Files\Created\WorkbookProperties.xlsx");
            new AdjustToContents().Create(@"C:\Excel Files\Created\AdjustToContents.xlsx");
            new HideUnhide().Create(@"C:\Excel Files\Created\HideUnhide.xlsx");
            new Outline().Create(@"C:\Excel Files\Created\Outline.xlsx");
            new Formulas().Create(@"C:\Excel Files\Created\Formulas.xlsx");
            new Collections().Create(@"C:\Excel Files\Created\Collections.xlsx");
            new NamedRanges().Create(@"C:\Excel Files\Created\NamedRanges.xlsx");
            new CopyingRanges().Create();
            new BlankCells().Create(@"C:\Excel Files\Created\BlankCells.xlsx");
            new TwoPages().Create(@"C:\Excel Files\Created\TwoPages.xlsx");
        }
    }
}
