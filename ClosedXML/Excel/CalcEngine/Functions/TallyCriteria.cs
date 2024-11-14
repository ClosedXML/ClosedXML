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
    /// A method to convert a value in the tally area to a number. If scalar value shouldn't be tallied, return null.
    /// </summary>
    private readonly Func<ScalarValue, double?> _toNumber;

    internal TallyCriteria()
        : this(static cellValue => cellValue.TryPickNumber(out var number) ? number : null)
    {
    }

    internal TallyCriteria(Func<ScalarValue, double?> toNumber)
    {
        _toNumber = toNumber;
    }

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
                var number = _toNumber(cellValue);
                if (number is not null)
                    state = state.Tally(number.Value);
            }
        }

        return state;
    }

    private static IEnumerable<XLSheetOffset> GetCombinedCoordinates(List<(XLSheetPoint Origin, IEnumerable<XLSheetPoint> Enumerable)> enumerables)
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
                // Do all enumerators have same offset?
                var allSame = true;
                var minOfs = GetOffset(0);
                for (var i = 1; i < enumerables.Count; ++i)
                {
                    var currentOfs = GetOffset(i);
                    var comparison = currentOfs.CompareTo(minOfs);
                    if (minOfs != currentOfs)
                        allSame = false;

                    if (comparison < 0)
                        minOfs = currentOfs;
                }

                // If all offsets are same, that means all criteria are
                // satisfied for same offset.
                if (allSame)
                    yield return minOfs;

                // Move all enumerators that point at the minimum offset
                // to the next element.
                for (var i = 0; i < enumerables.Count; ++i)
                {
                    var currentOfs = GetOffset(i);
                    if (currentOfs.CompareTo(minOfs) <= 0)
                    {
                        if (!enumerators[i].MoveNext())
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

        XLSheetOffset GetOffset(int i)
        {
            var origin = enumerables[i].Origin;
            var point = enumerators[i].Current;
            return point - origin;
        }
    }
}
