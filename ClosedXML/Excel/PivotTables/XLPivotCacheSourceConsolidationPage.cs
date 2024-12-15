using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// A page filter for pivot table that uses <see cref="XLPivotSourceConsolidation"/> as the source
/// of data. It is basically a container of strings that are displayed in a page filter above
/// the pivot table.
/// </summary>
internal class XLPivotCacheSourceConsolidationPage
{
    internal XLPivotCacheSourceConsolidationPage(List<string> pageItems)
    {
        PageItems = pageItems;
    }

    /// <summary>
    /// Page items (=names) displayed in the filter. The value is referenced
    /// through index by <see cref="XLPivotCacheSourceConsolidationRangeSet.Indexes"/>.
    /// </summary>
    internal IReadOnlyList<string> PageItems { get; }
}
