using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes uses a union of multiple scenarios in the workbook to
/// create data.
/// </summary>
internal sealed class XLPivotSourceConsolidation : IXLPivotSource
{
    /// <summary>
    /// Will application automatically create additional page filter in addition to the <see cref="Pages"/>.
    /// </summary>
    internal bool AutoPage { get; init; }

    /// <summary>
    /// <para>
    /// Custom page filters that toggle whether to display data from a particular
    /// <see cref="XLPivotCacheSourceConsolidationRangeSet">range set</see>.
    /// There can be 0..4 page filters. Each can have a different combination
    /// of range sets.
    /// </para>
    /// <para>
    /// Example:
    /// <example>
    /// The range sets are months and one page is <em>Q1</em>,<em>Q2</em>,<em>Q3</em>,<em>Q4</em>
    /// and second page filter is <em>Last month of quarter</em> and <em>Other months</em>. These
    /// page items are referenced by <see cref="XLPivotCacheSourceConsolidationRangeSet.Indexes"/>.
    /// </example>
    /// </para>
    /// </summary>
    public IReadOnlyList<XLPivotCacheSourceConsolidationPage> Pages { get; init; } = Array.Empty<XLPivotCacheSourceConsolidationPage>();

    /// <summary>
    /// Range sets that consists the cache source.
    /// </summary>
    public IReadOnlyList<XLPivotCacheSourceConsolidationRangeSet> RangeSets { get; init; } = Array.Empty<XLPivotCacheSourceConsolidationRangeSet>();

    public bool Equals(IXLPivotSource otherSource)
    {
        var other = otherSource as XLPivotSourceConsolidation;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        // This source should likely never be unified, so when there are two instances, mark them as different.
        return false;
    }

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out XLSheetRange? sheetArea)
    {
        throw new NotImplementedException("Consolidation pivot cache data source is not supported.");
    }
}
