using System;

namespace ClosedXML.Excel;

/// <summary>
/// A fluent API interface for adding another values to a <see cref="XLFilterType.Regular"/>
/// filter. It is chained by <see cref="IXLFilterColumn.AddFilter"/> method or
/// <see cref="IXLFilterColumn.AddDateGroupFilter"/>.
/// </summary>
public interface IXLFilteredColumn
{
    /// <summary>
    /// Add another value to a subset of allowed values for a <see cref="XLFilterType.Regular"/>
    /// filter. See <see cref="IXLFilterColumn.AddFilter"/> for more details.
    /// </summary>
    /// <param name="value">Value of the filter. The type is <c>XLCellValue</c>, but that's for
    /// convenience sake. The value is converted to a string and filter works with string.</param>
    /// <param name="reapply">Should the autofilter be immediately reapplied?</param>
    /// <returns>Fluent API allowing to add additional filter value.</returns>
    IXLFilteredColumn AddFilter(XLCellValue value, bool reapply = true);

    /// <summary>
    /// Add another grouping to a set of allowed groupings. See <see cref="IXLFilterColumn.AddDateGroupFilter"/>
    /// for more details.
    /// </summary>
    /// <returns>Fluent API allowing to add additional date group filter.</returns>
    IXLFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping, bool reapply = true);
}
