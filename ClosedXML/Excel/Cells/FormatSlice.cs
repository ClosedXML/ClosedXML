using System;
using System.Collections.Generic;
using ClosedXML.Excel.Formatting;
using ClosedXML.Utils;

namespace ClosedXML.Excel;

internal class FormatSlice : ISlice
{
    private readonly Slice<SliceValue> _slice = new();

    public bool IsEmpty => _slice.IsEmpty;

    public int MaxColumn => _slice.MaxColumn;

    public int MaxRow => _slice.MaxRow;

    public Dictionary<int, int>.KeyCollection UsedColumns => _slice.UsedColumns;

    public IEnumerable<int> UsedRows => _slice.UsedRows;

    public void Clear(XLSheetRange range)
    {
        _slice.Clear(range);
    }

    public void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete)
    {
        _slice.DeleteAreaAndShiftLeft(rangeToDelete);
    }

    public void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete)
    {
        _slice.DeleteAreaAndShiftUp(rangeToDelete);
    }

    public IEnumerator<XLSheetPoint> GetEnumerator(XLSheetRange range, bool reverse = false)
    {
        return _slice.GetEnumerator(range, reverse);
    }

    public void InsertAreaAndShiftDown(XLSheetRange range)
    {
        _slice.InsertAreaAndShiftDown(range);
    }

    public void InsertAreaAndShiftRight(XLSheetRange range)
    {
        _slice.InsertAreaAndShiftRight(range);
    }

    public bool IsUsed(XLSheetPoint address)
    {
        return _slice.IsUsed(address);
    }

    public void Swap(XLSheetPoint sp1, XLSheetPoint sp2)
    {
        _slice.Swap(sp1, sp2);
    }

    public void Set(XLSheetPoint point, XLStyleValue value)
    {
        var modified = _slice[point] with { StyleValue = value };
        _slice.Set(point, modified);
    }

    public void Set(XLSheetPoint point, XLCellFormatValue? value)
    {
        var modified = _slice[point] with { Format = value };
        _slice.Set(point, modified);
    }

    internal XLStyleValue? GetStyleValue(XLSheetPoint point)
    {
        return _slice[point].StyleValue;
    }

    internal XLCellFormatValue? GetFormat(XLSheetPoint point)
    {
        return _slice[point].Format;
    }

    /// <summary>
    /// Apply a deterministic format change. Deterministic = when inputs are equal, outputs must
    /// also be equal. That is needed to cache format modifications. This method doesn't register
    /// formats in the workbook styles, it only sets values in the slice.
    /// </summary>
    /// <param name="area">Area that should be modified.</param>
    /// <param name="modification">A deterministic modification.</param>
    /// <param name="resolver">A provider of format for non-materialized cells (e.g. column has a format and thus non-materialized cells should use column format).</param>
    internal void ApplyDeterministic(XLSheetRange area, Func<XLCellFormatValue, XLCellFormatValue> modification, Func<XLSheetPoint, XLCellFormatValue> resolver)
    {
        var cache = new Dictionary<XLCellFormatValue, XLCellFormatValue>(ReferenceEqualityComparer<XLCellFormatValue>.Instance);
        foreach (var point in area)
        {
            var format = GetFormat(point) ?? resolver(point);
            if (!cache.TryGetValue(format, out var modifiedFormat))
            {
                modifiedFormat = modification(format);
                cache.Add(format, modifiedFormat);
            }

            Set(point, modifiedFormat);
        }
    }

    private readonly record struct SliceValue(XLStyleValue? StyleValue, XLCellFormatValue? Format);
}
