using System;

namespace ClosedXML.Excel
{
    public interface IXLSheetProtection
    {
        Boolean Protected { get; set; }

        Boolean AutoFilter { get; set; }
        Boolean DeleteColumns { get; set; }
        Boolean DeleteRows { get; set; }
        Boolean FormatCells { get; set; }
        Boolean FormatColumns { get; set; }
        Boolean FormatRows { get; set; }
        Boolean InsertColumns { get; set; }
        Boolean InsertHyperlinks { get; set; }
        Boolean InsertRows { get; set; }
        Boolean Objects { get; set; }
        Boolean PivotTables { get; set; }
        Boolean Scenarios { get; set; }
        Boolean SelectLockedCells { get; set; }
        Boolean SelectUnlockedCells { get; set; }
        Boolean Sort { get; set; }

        IXLSheetProtection SetAutoFilter(); IXLSheetProtection SetAutoFilter(Boolean value);
        IXLSheetProtection SetDeleteColumns(); IXLSheetProtection SetDeleteColumns(Boolean value);
        IXLSheetProtection SetDeleteRows(); IXLSheetProtection SetDeleteRows(Boolean value);
        IXLSheetProtection SetFormatCells(); IXLSheetProtection SetFormatCells(Boolean value);
        IXLSheetProtection SetFormatColumns(); IXLSheetProtection SetFormatColumns(Boolean value);
        IXLSheetProtection SetFormatRows(); IXLSheetProtection SetFormatRows(Boolean value);
        IXLSheetProtection SetInsertColumns(); IXLSheetProtection SetInsertColumns(Boolean value);
        IXLSheetProtection SetInsertHyperlinks(); IXLSheetProtection SetInsertHyperlinks(Boolean value);
        IXLSheetProtection SetInsertRows(); IXLSheetProtection SetInsertRows(Boolean value);
        IXLSheetProtection SetObjects(); IXLSheetProtection SetObjects(Boolean value);
        IXLSheetProtection SetPivotTables(); IXLSheetProtection SetPivotTables(Boolean value);
        IXLSheetProtection SetScenarios(); IXLSheetProtection SetScenarios(Boolean value);
        IXLSheetProtection SetSelectLockedCells(); IXLSheetProtection SetSelectLockedCells(Boolean value);
        IXLSheetProtection SetSelectUnlockedCells(); IXLSheetProtection SetSelectUnlockedCells(Boolean value);
        IXLSheetProtection SetSort(); IXLSheetProtection SetSort(Boolean value);

        IXLSheetProtection Protect();
        IXLSheetProtection Protect(String password);
        IXLSheetProtection Unprotect();
        IXLSheetProtection Unprotect(String password);
    }
}
