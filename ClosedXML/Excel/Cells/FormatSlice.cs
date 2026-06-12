using System.Collections.Generic;
using ClosedXML.Excel.Formatting;

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

    public IEnumerator<Point> GetEnumerator(XLSheetRange range, bool reverse = false)
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

    public bool IsUsed(Point address)
    {
        return _slice.IsUsed(address);
    }

    public void Swap(Point sp1, Point sp2)
    {
        _slice.Swap(sp1, sp2);
    }

    public void Set(Point point, XLCellFormatValue? value)
    {
        var modified = _slice[point] with { Format = value };
        _slice.Set(point, modified);
    }

    internal void SetAll(XLSheetRange area, XLCellFormatValue? value)
    {
        _slice.SetAll(area, new SliceValue { Format = value });
    }

    internal XLCellFormatValue? GetFormat(Point point)
    {
        return _slice[point].Format;
    }

    // TODO Styles: FormatSlice should keep track of used format values so we don't have to go over all of them.
    internal void AddUsedFormat(HashSet<XLCellFormatValue> usedCellFormats)
    {
        var enumerator = GetEnumerator(XLSheetRange.Full);
        while (enumerator.MoveNext())
        {
            if (_slice[enumerator.Current].Format is { } format)
                usedCellFormats.Add(format);
        }
    }

    private readonly record struct SliceValue(XLCellFormatValue? Format);
}
