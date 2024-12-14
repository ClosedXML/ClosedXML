using System;

namespace ClosedXML.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes data from a connection
/// to external source of data (e.g. database or a workbook).
/// </summary>
internal sealed class XLPivotSourceConnection : IXLPivotSource
{
    public XLPivotSourceConnection(uint connectionId)
    {
        ConnectionId = connectionId;
    }

    public uint ConnectionId { get; }

    public bool Equals(IXLPivotSource otherSource)
    {
        var other = otherSource as XLPivotSourceConnection;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return ConnectionId == other.ConnectionId;
    }

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out XLSheetRange? sheetArea)
    {
        throw new NotImplementedException("Pivot cache source using a connection is not supported.");
    }
}
