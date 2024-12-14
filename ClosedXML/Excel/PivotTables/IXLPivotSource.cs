using System;

namespace ClosedXML.Excel;

/// <summary>
/// An abstraction of source data for a <see cref="XLPivotCache"/>. Implementations must correctly
/// implement equals.
/// </summary>
internal interface IXLPivotSource : IEquatable<IXLPivotSource>
{
    /// <summary>
    /// Try to determine actual area of the source reference in the
    /// workbook. Source reference might not be valid in the workbook, some might
    /// not be supported.
    /// </summary>
    bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out XLSheetRange? sheetArea);
}
