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
    class Program
    {
        static void Main(string[] args)
        {
            new HelloWorld().Create(@"c:\HelloWorld.xlsx");
            new BasicTable().Create(@"c:\BasicTable.xlsx");
            new StyleExamples().Create();
            new ColumnSettings().Create(@"c:\ColumnSettings.xlsx");
            new RowSettings().Create(@"c:\RowSettings.xlsx");
            new MergeCells().Create(@"c:\MergedCells.xlsx");
            new InsertRows().Create(@"c:\InsertRows.xlsx");
            new InsertColumns().Create(@"c:\InsertColumns.xlsx");
            new ColumnCollection().Create(@"c:\ColumnCollection.xlsx");
            new DataTypes().Create(@"c:\DataTypes.xlsx");
            new MultipleSheets().Create(@"c:\MultipleSheets.xlsx");
            new RowCollection().Create(@"c:\RowCollection.xlsx");
            new DefiningRanges().Create(@"c:\DefiningRanges.xlsx");
            new ClearingRanges().Create(@"c:\ClearingRanges.xlsx");
            new DeletingRanges().Create(@"c:\DeletingRanges.xlsx");
            new Margins().Create(@"c:\Margins.xlsx");
            new Page().Create(@"c:\Page.xlsx");
            new HeaderFooters().Create(@"c:\HeaderFooters.xlsx");
            new Sheets().Create(@"c:\Sheets.xlsx");
            new SheetTab().Create(@"c:\SheetTab.xlsx");
        }
    }
}