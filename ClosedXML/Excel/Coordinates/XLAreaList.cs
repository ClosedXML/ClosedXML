using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// An immutable collection of areas. An equivalent of <c>ST_Sqref</c> (sequence of references).
/// List doesn't allow duplicate areas.
/// </summary>
internal class XLAreaList : IEnumerable<XLSheetRange>
{
    internal static readonly XLAreaList Empty = new(new List<XLSheetRange>());
    private readonly List<XLSheetRange> _areas;

    internal XLAreaList(List<XLSheetRange> areas)
    {
        _areas = areas;
    }

    internal int Count => _areas.Count;

    internal XLAreaList With(XLSheetRange area)
    {
        if (_areas.Contains(area))
            return this;

        return new XLAreaList(new List<XLSheetRange>(_areas) { area });
    }

    internal XLAreaList Without(XLSheetRange area)
    {
        var indexToDelete = _areas.IndexOf(area);
        if (indexToDelete == -1)
            return this;

        var newList = new List<XLSheetRange>(_areas);
        newList.RemoveAt(indexToDelete);
        return new XLAreaList(newList);
    }

    internal XLAreaList InsertAndShiftDown(XLSheetRange insertedArea)
    {
        var groove = insertedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<XLSheetRange>(_areas.Count);
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
                var touchingArea = insertedArea.Intersect(new XLSheetRange(XLHelper.MinRowNumber, originalArea.LeftColumn, XLHelper.MaxRowNumber, originalArea.RightColumn));
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

    internal XLAreaList InsertAndShiftRight(XLSheetRange insertedArea)
    {
        // Find an area that will be shifted (if there is any). It must be in a groove
        // from insertedArea to the right end of a sheet
        var groove = insertedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<XLSheetRange>(_areas.Count);
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
                var touchingArea = insertedArea.Intersect(new XLSheetRange(originalArea.TopRow, XLHelper.MinColumnNumber, originalArea.BottomRow, XLHelper.MaxColumnNumber));
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

    internal XLAreaList DeleteAndShiftUp(XLSheetRange deletedArea)
    {
        var groove = deletedArea.ExtendBelow(XLHelper.MaxRowNumber);
        var result = new List<XLSheetRange>(_areas.Count);
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

    internal XLAreaList DeleteAndShiftLeft(XLSheetRange deletedArea)
    {
        var groove = deletedArea.ExtendRight(XLHelper.MaxColumnNumber);
        var result = new List<XLSheetRange>(_areas.Count);
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

    public IEnumerator<XLSheetRange> GetEnumerator()
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
}
