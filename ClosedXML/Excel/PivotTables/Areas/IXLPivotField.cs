#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLSubtotalFunction
    {
        Automatic,
        Sum,
        Count,
        Average,
        Maximum,
        Minimum,
        Product,
        CountNumbers,
        StandardDeviation,
        PopulationStandardDeviation,
        Variance,
        PopulationVariance,
    }

    public enum XLPivotLayout { Outline, Tabular, Compact }

    /// <summary>
    /// A fluent API representation of a field on an <see cref="IXLPivotTable.RowLabels">row</see>,
    /// <see cref="IXLPivotTable.ColumnLabels">column</see> or <see cref="IXLPivotTable.ReportFilters">
    /// filter</see> axis of a <see cref="IXLPivotTable"/>.
    /// </summary>
    /// <remarks>
    /// If the field is a 'data' field, a lot of properties don't make sense and can't be set. In
    /// such case, the setter will throw <exception cref="InvalidOperationException"/> and getter
    /// will return default value for the field.
    /// </remarks>
    public interface IXLPivotField
    {
        /// <summary>
        /// <para>
        /// Name of the field in a pivot table <see cref="IXLPivotTable.PivotCache"/>. If the field
        /// is 'data' field, return <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/>.
        /// </para>
        /// <para>
        /// Note that field name in pivot cache is generally same as in the source data range, but
        /// not always. Field names are unique in the cache and if the source data range contains
        /// duplicate column names, the cache will rename them to keep all names unique.
        /// </para>
        /// </summary>
        String SourceName { get; }

        /// <summary>
        /// <see cref="CustomName"/> of the field in the pivot table. Custom name is a unique
        /// across all fields used in the pivot table (e.g. if same field is added to values area
        /// multiple times, it must have custom name, e.g. <c>Sum1 of Field</c>,
        /// <c>Sum2 of Field</c>).
        /// </summary>
        /// <exception cref="ArgumentException">When setting name to a name that is already used by
        ///     another field.</exception>
        String CustomName { get; set; }

        String SubtotalCaption { get; set; }
        IReadOnlyCollection<XLSubtotalFunction> Subtotals { get; }
        Boolean IncludeNewItemsInFilter { get; set; }
        Boolean Outline { get; set; }
        Boolean Compact { get; set; }
        Boolean? SubtotalsAtTop { get; set; }
        Boolean RepeatItemLabels { get; set; }
        Boolean InsertBlankLines { get; set; }
        Boolean ShowBlankItems { get; set; }
        Boolean InsertPageBreaks { get; set; }

        /// <summary>
        /// Are all items of the field collapsed?
        /// </summary>
        /// <remarks>
        /// If only a subset of items is collapsed, getter returns <c>false</c>.
        /// </remarks>
        Boolean Collapsed { get; set; }

        XLPivotSortType SortType { get; set; }

        /// <inheritdoc cref="CustomName"/>
        IXLPivotField SetCustomName(String value);

        IXLPivotField SetSubtotalCaption(String value);

        IXLPivotField AddSubtotal(XLSubtotalFunction value);

        IXLPivotField SetIncludeNewItemsInFilter(Boolean value = true);

        IXLPivotField SetLayout(XLPivotLayout value);

        IXLPivotField SetSubtotalsAtTop(Boolean value = true);

        IXLPivotField SetRepeatItemLabels(Boolean value = true);

        IXLPivotField SetInsertBlankLines(Boolean value = true);

        IXLPivotField SetShowBlankItems(Boolean value = true);

        IXLPivotField SetInsertPageBreaks(Boolean value = true);

        IXLPivotField SetCollapsed(Boolean value = true);

        IXLPivotField SetSort(XLPivotSortType value);

        /// <summary>
        /// Selected values for <see cref="IXLPivotTable.ReportFilters"/> filter of the pivot
        /// table. Empty for non-filter fields.
        /// </summary>
        IReadOnlyList<XLCellValue> SelectedValues { get; }

        /// <summary>
        /// Add a value to selected values of a filter field (<see cref="IXLPivotTable.ReportFilters"/>).
        /// Doesn't do anything, if this field is not a filter fields.
        /// </summary>
        IXLPivotField AddSelectedValue(XLCellValue value);

        /// <summary>
        /// Add a values to a selected values of a filter field. Doesn't do anything if this field
        /// is not a filter fields.
        /// </summary>
        IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values);

        IXLPivotFieldStyleFormats StyleFormats { get; }

        Boolean IsOnRowAxis { get; }
        Boolean IsOnColumnAxis { get; }
        Boolean IsInFilterList { get; }

        /// <summary>
        /// Index of a field in <see cref="XLPivotTable.PivotFields">all pivot fields</see> or -2
        /// for data field.
        /// </summary>
        Int32 Offset { get; }
    }
}
