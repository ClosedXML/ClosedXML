using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLSheetProtectionElements
    {
        None = 0,
        AutoFilter = 1 << 1,
        DeleteColumns = 1 << 2,
        DeleteRows = 1 << 3,
        EditObjects = 1 << 4,
        EditScenarios = 1 << 5,
        FormatCells = 1 << 6,
        FormatColumns = 1 << 7,
        FormatRows = 1 << 8,
        InsertColumns = 1 << 9,
        InsertHyperlinks = 1 << 10,
        InsertRows = 1 << 11,
        PivotTables = 1 << 12,
        SelectLockedCells = 1 << 13,
        SelectUnlockedCells = 1 << 14,
        Sort = 1 << 15,

        DeleteEverything = DeleteColumns | DeleteRows,
        FormatEverything = FormatCells | FormatColumns | FormatRows,
        InsertEverything = InsertColumns | InsertHyperlinks | InsertRows,
        SelectEverything = SelectLockedCells | SelectUnlockedCells,

        Everything = AutoFilter
            | DeleteColumns | DeleteRows
            | EditObjects | EditScenarios
            | FormatCells | FormatColumns | FormatRows
            | InsertColumns | InsertHyperlinks | InsertRows
            | PivotTables
            | SelectLockedCells | SelectUnlockedCells
            | Sort
    }
}
