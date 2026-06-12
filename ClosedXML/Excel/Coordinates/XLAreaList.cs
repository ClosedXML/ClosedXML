using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel;

/// <summary>
/// An immutable collection of areas. An equivalent of <c>ST_Sqref</c> (sequence of references).
/// List doesn't allow duplicate areas.
/// </summary>
internal class XLAreaList : IEnumerable<Area>
{
    internal static readonly XLAreaList Empty = new(new List<Area>());
    private readonly List<Area> _areas;

    internal XLAreaList(Area area)
    {
        _areas = new List<Area>(1) { area };
    }

    internal XLAreaList(List<Area> areas)
    {
        _areas = areas;
    }

    internal int Count => _areas.Count;

    internal Area this[int idx] => _areas[idx];

    internal static XLAreaList FromRange(XLWorksheet worksheet, IXLRange range)
    {
        ThrowOnDifferentSheet(worksheet, range);
        return new XLAreaList(Area.FromRangeAddress(range.RangeAddress));
    }

    /// <exception cref="ArgumentException">Sequence is empty or a range is from a different sheet.</exception>
    internal static XLAreaList FromRanges(XLWorksheet worksheet, IEnumerable<IXLRange> value)
    {
        var areas = new List<Area>();
        foreach (var range in value)
        {
            ThrowOnDifferentSheet(worksheet, range);
            areas.Add(Area.FromRangeAddress(range.RangeAddress));
        }

        if (areas.Count == 0)
            throw new ArgumentException("Sequence is empty. At least one range is required.");

        return new XLAreaList(areas);
    }

    internal XLAreaList With(Area area)
    {
        if (_areas.Contains(area))
            return this;

        return new XLAreaList(new List<Area>(_areas) { area });
    }

    internal XLAreaList Without(Area area)
    {
        var indexToDelete = _areas.IndexOf(area);
        if (indexToDelete == -1)
            return this;

        var newList = new List<Area>(_areas);
        newList.RemoveAt(indexToDelete);
        return new XLAreaList(newList);
    }

    internal XLAreaList InsertAndShiftDown(Area insertedArea)
    {
        var groove = insertedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }

            var insertWontSplitOriginalArea = insertedArea.LeftColumn <= originalArea.LeftColumn && insertedArea.RightColumn >= originalArea.RightColumn;
            if (insertWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrExtendDown(insertedArea.TopRow, insertedArea.Height);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else if (insertedArea.TopRow == originalArea.BottomRow + 1)
            {
                result.Add(originalArea);
                var touchingArea = insertedArea.Intersect(new Area(XLHelper.MinRowNumber, originalArea.LeftColumn, XLHelper.MaxRowNumber, originalArea.RightColumn));
                if (touchingArea is not null)
                    result.Add(touchingArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is null)
                    continue;

                // There is something to shift, so shift it downwards
                var shiftedArea = inGrooveArea.Value.ShiftOrExtendDown(insertedArea.TopRow, insertedArea.Height);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList InsertAndShiftRight(Area insertedArea)
    {
        // Find an area that will be shifted (if there is any). It must be in a groove
        // from insertedArea to the right end of a sheet
        var groove = insertedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            var insertWontSplitOriginalArea = insertedArea.TopRow <= originalArea.TopRow && insertedArea.BottomRow >= originalArea.BottomRow;
            if (insertWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrExtendRight(insertedArea.LeftColumn, insertedArea.Width);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else if (insertedArea.LeftColumn == originalArea.RightColumn + 1)
            {
                result.Add(originalArea);
                var touchingArea = insertedArea.Intersect(new Area(originalArea.TopRow, XLHelper.MinColumnNumber, originalArea.BottomRow, XLHelper.MaxColumnNumber));
                if (touchingArea is not null)
                    result.Add(touchingArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is null)
                    continue;

                // There is something to shift, so shift it rightwards
                var shiftedArea = inGrooveArea.Value.ShiftOrExtendRight(insertedArea.LeftColumn, insertedArea.Width);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftUp(Area deletedArea)
    {
        var groove = deletedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }
            var deleteWontSplitOriginalArea = deletedArea.LeftColumn <= originalArea.LeftColumn && deletedArea.RightColumn >= originalArea.RightColumn;
            if (deleteWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrShrinkUp(deletedArea.TopRow, deletedArea.Height);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it upwards
                    var shiftedArea = inGrooveArea.Value.ShiftOrShrinkUp(deletedArea.TopRow, deletedArea.Height);
                    if (shiftedArea is not null)
                        result.Add(shiftedArea.Value);
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftLeft(Area deletedArea)
    {
        var groove = deletedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            var deleteWontSplitOriginalArea = deletedArea.TopRow <= originalArea.TopRow && deletedArea.BottomRow >= originalArea.BottomRow;
            if (deleteWontSplitOriginalArea)
            {
                var shiftedArea = originalArea.ShiftOrShrinkLeft(deletedArea.LeftColumn, deletedArea.Width);
                if (shiftedArea is not null)
                    result.Add(shiftedArea.Value);
            }
            else
            {
                var inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it leftward
                    var shiftedArea = inGrooveArea.Value.ShiftOrShrinkLeft(deletedArea.LeftColumn, deletedArea.Width);
                    if (shiftedArea is not null)
                        result.Add(shiftedArea.Value);
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteWithoutShift(Area deletedArea)
    {
        var result = new List<Area>(_areas.Count);
        foreach (var originalArea in _areas)
            originalArea.Exclude(deletedArea, result);

        return new XLAreaList(result);
    }

    internal XLAreaList GetConsolidated()
    {
        return XLRangeConsolidationEngine.Consolidate(this);
    }

    internal bool IntersectsWith(Area otherArea)
    {
        foreach (var area in _areas)
        {
            if (area.Intersects(otherArea))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Return areas in the list (with the original size) intersecting with the <paramref name="otherArea"/>.
    /// </summary>
    internal IEnumerable<Area> IntersectingWith(Area otherArea)
    {
        foreach (var area in _areas)
        {
            if (area.Intersects(otherArea))
                yield return area;
        }
    }

    /// <summary>
    /// A helper function used mostly in copy&amp;paste functionality. It takes the areas,
    /// intersects them with the <paramref name="areaToCopy"/> and shifts it to the <paramref name="target"/>.
    /// If there are areas, return it in the <paramref name="result"/>.
    /// </summary>
    internal bool TryCopyAreaTo(Point target, Area areaToCopy, [NotNullWhen(true)] out XLAreaList? result)
    {
        var rowShift = target.Row - areaToCopy.FirstPoint.Row;
        var columnShift = target.Column - areaToCopy.FirstPoint.Column;
        List<Area>? copyList = null;
        foreach (var area in _areas)
        {
            if (area.Intersect(areaToCopy) is not { } intersection)
                continue;

            // End can but cut off, but the area will always have at least 1x1 so it is valid
            if (intersection.ShiftAndClip(rowShift, columnShift) is not { } shiftedArea)
                continue;

            copyList ??= new List<Area>();
            copyList.Add(shiftedArea);
        }

        if (copyList is not null)
        {
            result = new XLAreaList(copyList);
            return true;
        }

        result = null;
        return false;
    }

    internal XLAreaList Excluding(Area excludedArea)
    {
        if (!IntersectsWith(excludedArea))
            return this;

        var list = new List<Area>();
        foreach (var area in _areas)
            area.Exclude(excludedArea, list);

        return new XLAreaList(list);
    }

    public IEnumerator<Area> GetEnumerator()
    {
        return _areas.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal string ToSpaceList()
    {
        return string.Join(" ", _areas);
    }

    private static void ThrowOnDifferentSheet(XLWorksheet worksheet, IXLRange range)
    {
        if (range.Worksheet is not null && range.Worksheet != worksheet)
            throw new ArgumentException($"Range {range} belongs to worksheet {range.Worksheet.Name}, but must be from worksheet {worksheet.Name}.");
    }
}
