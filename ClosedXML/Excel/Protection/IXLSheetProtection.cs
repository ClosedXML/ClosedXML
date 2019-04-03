// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

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

    public interface IXLSheetProtection : ICloneable
    {
        Algorithm Algorithm { get; }

        XLSheetProtectionElements AllowedElements { get; set; }

        Boolean IsProtected { get; }

        /// <summary>
        /// Adds the sheet protection element to the list of allowed elements.
        /// Beware that if you pass through <see cref="XLSheetProtectionElements.None" />, this will have no effect.
        /// </summary>
        /// <param name="element">The sheet protection element to add</param>
        /// <param name="allowed">Set to <c>true</c> to allow the element or <c>false</c> to disallow the element</param>
        /// <returns>The current sheet protection</returns>
        IXLSheetProtection AllowElement(XLSheetProtectionElements element, Boolean allowed = true);

        IXLSheetProtection AllowEverything();

        IXLSheetProtection AllowNone();

        IXLSheetProtection CopyFrom(IXLSheetProtection sheetProtection);

        /// <summary>
        /// Removes the sheet protection element to the list of allowed elements.
        /// Beware that if you pass through <see cref="XLSheetProtectionElements.None" />, this will have no effect.
        /// </summary>
        /// <param name="element">The sheet protection element to remove</param>
        /// <returns>The current sheet protection</returns>
        IXLSheetProtection DisallowElement(XLSheetProtectionElements element);

        IXLSheetProtection Protect();

        IXLSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm);

        IXLSheetProtection Unprotect();

        IXLSheetProtection Unprotect(String password);
    }
}
