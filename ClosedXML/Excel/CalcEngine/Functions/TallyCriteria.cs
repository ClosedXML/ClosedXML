using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine.Functions;

/// <summary>
/// Tally for <c>{SUM,COUNT,AVERAGE}IF/S</c> and database function. The created tally must contain
/// all selection areas and associated criteria. The main <see cref="Tally{T}"/> function is then
/// called with values that will be tallied, based on the areas+criteria in the tally object.
/// </summary>
internal class TallyCriteria : ITally
{
    /// <summary>
    /// A collection of areas that are tested and if all satisfy the criteria, corresponding values
    /// in the tally areas are tallied.
    /// </summary>
    private readonly List<(XLRangeAddress Area, Criteria Criteria)> _criteriaRanges = new();

    /// <summary>
    /// Add criteria to the tally that limit which values should be tallied.
    /// </summary>
    internal void Add(XLRangeAddress area, Criteria criteria)
    {
        _criteriaRanges.Add((area, criteria));
    }

    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        // All criteria functions permit only area reference arguments. Excel ensures this
        // invariant by grammar, we just check the the argument value.
        var talliedAreas = new List<XLRangeAddress>(args.Length);
        foreach (var arg in args)
        {
            if (!arg.TryPickArea(out var tallyArea, out var error))
                return error;

            talliedAreas.Add(tallyArea);
        }

        // For each selection area and its criteria, get list of points that satisfy the criteria.
        var criteriaPoints = new List<(XLSheetPoint Origin, IEnumerable<XLSheetPoint> Enumerable)>();
        foreach (var (area, criteria) in _criteriaRanges)
        {
            // This is a lazy IEnumerable, it's not yet evaluated.
            var areaCriteriaPoints = ctx.GetCriteriaPoints(area, criteria);
            var origin = XLSheetRange.FromRangeAddress(area).FirstPoint;
            criteriaPoints.Add((origin, areaCriteriaPoints));
        }

        // Get list of points that satisfy all criteria
        var talliedCoordinates = GetCombinedCoordinates(criteriaPoints);

        var state = initialState;
        foreach (var (rowOfs, colOfs) in talliedCoordinates)
        {
            foreach (var area in talliedAreas)
            {
                var origin = area.FirstAddress;
                var shifted = new XLSheetPoint(origin.RowNumber + rowOfs, origin.ColumnNumber + colOfs);
                var cellValue = ctx.GetCellValue(area.Worksheet, shifted.Row, shifted.Column);
                if (cellValue.TryPickNumber(out var number))
                    state = state.Tally(number);
            }
        }

        return state;
    }

    private static IEnumerable<(int RowOfs, int ColOfs)> GetCombinedCoordinates(List<(XLSheetPoint Origin, IEnumerable<XLSheetPoint> Enumerable)> enumerables)
    {
        var enumerators = enumerables.Select(e => e.Enumerable.GetEnumerator()).ToList();
        try
        {
            // Move to the first element
            foreach (var enumerator in enumerators)
            {
                if (!enumerator.MoveNext())
                    yield break;
            }

            // Until all elements are processed.
            while (true)
            {
                // Get minimum point from all enumerators
                var min = enumerators[0].Current;
                var minOrigin = enumerables[0].Origin;
                for (var i = 1; i < enumerables.Count; ++i)
                {
                    var current = enumerators[i].Current;
                    var comparison = current.CompareTo(min);
                    if (comparison < 0)
                    {
                        min = current;
                        minOrigin = enumerables[i].Origin;
                    }
                }

                // Returns the offset of the minimum point
                yield return (min.Row - minOrigin.Row, min.Column - minOrigin.Column);

                // Move all enumerators that point at the minimum to the next element
                foreach (var enumerator in enumerators)
                {
                    // If would likely suffice, because enumerator doesn't contain duplicates
                    // and it moves one element at a time.
                    while (enumerator.Current.CompareTo(min) <= 0)
                    {
                        if (!enumerator.MoveNext())
                            yield break;
                    }
                }
            }
        }
        finally
        {
            foreach (var enumerator in enumerators)
                enumerator.Dispose();
        }
    }
}
