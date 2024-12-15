using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel;

/// <summary>
/// One of ranges that form a <see cref="XLPivotSourceConsolidation"/> source for a <see cref="XLPivotCache"/>.
/// </summary>
internal class XLPivotCacheSourceConsolidationRangeSet
{
    /// <summary>
    /// Indexes into the <see cref="XLPivotSourceConsolidation.Pages"/>. If the value is null
    /// and page filter exists, it is displayed as a blank. There can be at most 4 indexes, because
    /// there can be at most 4 page filters.
    /// </summary>
    public IReadOnlyList<uint?> Indexes { get; init; } = Array.Empty<uint?>();

    /// <summary>
    /// If range set is from another workbook, a relationship id to the workbook from cache definition.
    /// </summary>
    internal string? RelId { get; init; }

    [MemberNotNullWhen(true, nameof(TableOrName))]
    [MemberNotNullWhen(false, nameof(Area))]
    internal bool UsesName => TableOrName is not null;

    internal string? TableOrName { get; init; }

    internal XLBookArea? Area { get; init; }
}
