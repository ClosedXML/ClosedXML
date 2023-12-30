using System;

namespace ClosedXML.Excel
{
    public enum XLTopBottomType { Items, Percent }

    public enum XLDateTimeGrouping { Year, Month, Day, Hour, Minute, Second }

    /// <summary>
    /// <para>
    /// AutoFilter filter configuration for one column in an autofilter <see cref="IXLAutoFilter.Range">area</see>.
    /// Filters determine visibility of rows in the autofilter area. Value in the row must satisfy
    /// all filters in all columns in order for row to be visible, otherwise it is <see cref="IXLRow.IsHidden"/>.
    /// </para>
    /// <para>
    /// Column can have only one type of filter, so it's not possible to combine several different
    /// filter types on one column. Methods for adding filters clear other types or remove
    /// previously set filters when needed. Some types of filters can have multiple conditions (e.g.
    /// <see cref="XLFilterType.Regular"/> can have many values while <see cref="XLFilterType.Dynamic"/>
    /// can be only one).
    /// </para>
    /// <para>
    /// Whenever filter configuration changes, the filters are immediately reapplied.
    /// </para>
    /// </summary>
    public interface IXLFilterColumn
    {
        /// <summary>
        /// Remove all filters from the column.
        /// </summary>
        /// <remarks>
        /// Does not reapply filters, visibility of rows isn't changed.
        /// </remarks>
        void Clear();

        /// <summary>
        /// <para>
        /// Switch to the <see cref="XLFilterType.Regular"/> filter if filter column has a
        /// different type (for current type <see cref="FilterType"/>) and add
        /// <paramref name="value"/> to a set of allowed values. Excel displays regular filter as
        /// a list of possible values in a column with checkbox next to it and user can check which
        /// one should be displayed.
        /// </para>
        /// <para>
        /// From technical perspective, the passed <paramref name="value"/> is converted to
        /// a localized string (using current locale) and the column values satisfy the filter
        /// condition, when the <see cref="IXLCell.GetFormattedString">formatted string of a cell
        /// </see> matches any filter string.
        /// </para>
        /// <para>
        /// Examples of less intuitive behavior: filter value is <c>2.5</c> in locale cs-CZ that
        /// uses "<em>,</em>" as a decimal separator. The passed <paramref name="value"/>
        /// is number 2.5, converted immediately to a string <em>2,5</em>. The string is used for
        /// comparison with values of cells in the column:
        /// <list type="bullet">
        ///  <item>Number 2.5 formatted with two decimal places as <em>2,50</em> will not match.</item>
        ///  <item>Number 2.5 with default formatting will be matched, because its string is
        ///        <em>2,5</em> in cs-CZ locale (but not in others, e.g. en-US locale).</item>
        ///  <item>Text <em>2,5</em> will be matched.</item>
        /// </list>
        /// </para>
        /// </summary>
        /// <remarks>
        /// This behavior of course highly depends on locale and working with same file on two
        /// different locales might lead to different results.
        /// </remarks>
        /// <param name="value">Value of the filter. The type is <c>XLCellValue</c>, but that's for
        /// convenience sake. The value is converted to a string and filter works with string.</param>
        /// <returns>Fluent API allowing to add additional filter value.</returns>
        IXLFilteredColumn AddFilter(XLCellValue value);

        /// <summary>
        /// <para>
        /// Enable autofilter (if needed), switch to the <see cref="XLFilterType.Regular"/> filter
        /// if filter column has a different type (for current type <see cref="FilterType"/>) and
        /// add a filter that is satisfied when cell value is a <see cref="XLDataType.DateTime"/>
        /// and the tested date has same components from <paramref name="dateTimeGrouping"/>
        /// component down to the <see cref="XLDateTimeGrouping.Second"/> component with same value
        /// as the <paramref name="dateTimeGrouping"/>.
        /// </para>
        /// <para>
        /// <example>
        /// Example:
        /// <code>
        /// // Filter will be satisfied if the cell value is a XLDataType.DateTime and the month,
        /// // day, hour, minute and second are same as the passed date.
        /// AddDateGroupFilter(new DateTime(2023, 7, 15), XLDateTimeGrouping.Month)
        /// </code>
        /// </example>
        /// </para>
        /// <para>
        /// There can be multiple date group filters and they are <see cref="XLFilterType.Regular"/>
        /// filter types, i.e. they don't delete filters from <see cref="AddFilter"/>. The cell
        /// value is satisfied, if it matches any of the text values from <see cref="AddFilter"/>
        /// or any date group filter.
        /// </para>
        /// </summary>
        /// <param name="date">Date which components are compared with date values of the column.</param>
        /// <param name="dateTimeGrouping">
        /// Starting component of the grouping. Tested date must match all date components of the
        /// <paramref name="date"/> from this one to the <see cref="XLDateTimeGrouping.Second"/>.
        /// </param>
        /// <returns>Fluent API allowing to add additional date time group value.</returns>
        IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping);

        void Top(Int32 value, XLTopBottomType type = XLTopBottomType.Items);

        void Bottom(Int32 value, XLTopBottomType type = XLTopBottomType.Items);

        void AboveAverage();

        void BelowAverage();

        IXLFilterConnector EqualTo(XLCellValue value);

        IXLFilterConnector NotEqualTo(XLCellValue value);

        IXLFilterConnector GreaterThan(XLCellValue value);

        IXLFilterConnector LessThan(XLCellValue value);

        IXLFilterConnector EqualOrGreaterThan(XLCellValue value);

        IXLFilterConnector EqualOrLessThan(XLCellValue value);

        void Between(XLCellValue minValue, XLCellValue maxValue);

        void NotBetween(XLCellValue minValue, XLCellValue maxValue);

        IXLFilterConnector BeginsWith(String value);

        IXLFilterConnector NotBeginsWith(String value);

        IXLFilterConnector EndsWith(String value);

        IXLFilterConnector NotEndsWith(String value);

        IXLFilterConnector Contains(String value);

        IXLFilterConnector NotContains(String value);

        /// <summary>
        /// Current filter type used by the filter columns.
        /// </summary>
        XLFilterType FilterType { get; }

        /// <summary>
        /// Configuration of a <see cref="XLFilterType.TopBottom"/> filter. It contains how many
        /// items/percent (depends on <see cref="TopBottomType"/>) should be filter accept.
        /// </summary>
        /// <remarks>
        /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XLFilterType.TopBottom"/>.
        /// </remarks>
        Int32 TopBottomValue { get; }

        /// <summary>
        /// Configuration of a <see cref="XLFilterType.TopBottom"/> filter. It contains the content
        /// interpretation of a <see cref="TopBottomValue"/> property, either how many items or how
        /// many percents.
        /// </summary>
        /// <remarks>
        /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XLFilterType.TopBottom"/>.
        /// </remarks>
        XLTopBottomType TopBottomType { get; }

        /// <summary>
        /// Configuration of a <see cref="XLFilterType.TopBottom"/> filter. It determines if filter
        /// should accept items from top or bottom.
        /// </summary>
        /// <remarks>
        /// Returns undefined value, if <see cref="FilterType"/> is not <see cref="XLFilterType.TopBottom"/>.
        /// </remarks>
        XLTopBottomPart TopBottomPart { get; }
        XLFilterDynamicType DynamicType { get; }
        Double DynamicValue { get; }

        IXLFilterColumn SetFilterType(XLFilterType value);

        IXLFilterColumn SetTopBottomValue(Int32 value);

        IXLFilterColumn SetTopBottomType(XLTopBottomType value);

        IXLFilterColumn SetTopBottomPart(XLTopBottomPart value);

        IXLFilterColumn SetDynamicType(XLFilterDynamicType value);

        IXLFilterColumn SetDynamicValue(Double value);
    }
}
