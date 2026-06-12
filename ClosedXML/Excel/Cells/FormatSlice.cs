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

    public void Clear(Area area)
    {
        _slice.Clear(area);
    }

    public void DeleteAreaAndShiftLeft(Area areaToDelete)
    {
        _slice.DeleteAreaAndShiftLeft(areaToDelete);
    }

    public void DeleteAreaAndShiftUp(Area areaToDelete)
    {
        _slice.DeleteAreaAndShiftUp(areaToDelete);
    }

    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false)
    {
        return _slice.GetEnumerator(area, reverse);
    }

    public void InsertAreaAndShiftDown(Area areaToInsert)
    {
        _slice.InsertAreaAndShiftDown(areaToInsert);
    }

    public void InsertAreaAndShiftRight(Area areaToInsert)
    {
        _slice.InsertAreaAndShiftRight(areaToInsert);
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

    internal void SetAll(Area area, XLCellFormatValue? value)
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
        var enumerator = GetEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            if (_slice[enumerator.Current].Format is { } format)
                usedCellFormats.Add(format);
        }
    }

    private readonly record struct SliceValue(XLCellFormatValue? Format);
}
