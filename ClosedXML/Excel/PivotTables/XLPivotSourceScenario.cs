using System;

namespace ClosedXML.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes uses scenarios in the workbook to
/// create data.
/// </summary>
internal sealed class XLPivotSourceScenario : IXLPivotSource
{
    public bool Equals(IXLPivotSource other)
    {
        return other is XLPivotSourceScenario;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXLPivotSource other && Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out XLSheetRange? sheetArea)
    {
        throw new NotImplementedException("Scenario pivot cache data source is not supported.");
    }
}
