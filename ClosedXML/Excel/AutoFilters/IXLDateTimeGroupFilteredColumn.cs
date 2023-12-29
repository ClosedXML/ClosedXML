using System;

namespace ClosedXML.Excel;

/// <summary>
/// A fluent API interface for adding another values to a <see cref="XLFilterType.Regular"/>
/// filter. It is chained by <see cref="IXLFilterColumn.AddDateGroupFilter"/> method.
/// </summary>
/// <para>
/// Whenever filter configuration changes, the filters are immediately reapplied.
/// </para>
public interface IXLDateTimeGroupFilteredColumn
{
    /// <summary>
    /// Add another grouping to a set of allowed groupings. See <see cref="IXLFilterColumn.AddDateGroupFilter"/>
    /// for more details.
    /// </summary>
    /// <returns>Fluent API allowing to add additional date group filter.</returns>
    IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping);
}
