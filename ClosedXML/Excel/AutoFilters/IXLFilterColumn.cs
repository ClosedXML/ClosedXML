using System;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    public enum XLTopBottomType { Items, Percent }

    public enum XLDateTimeGrouping { Year, Month, Day, Hour, Minute, Second }

    /// <summary>
    /// <para>
    /// AutoFilter filter configuration for one column in an autofilter <see cref="IXLAutoFilter.Range">area</see>.
    /// Filters determine visibility of rows in the autofilter area. Column can have multiple
    /// filters, each specifying a different condition. Value in the row must satisfy all filters
    /// in order for row to be visible.
    /// </para>
    /// <para>
    /// Column can have only one type of filter, so it's not possible to combine several different
    /// filter types on one column. Methods for adding filters clear other types or remove
    /// previously set filters when needed. Some types of filters can have multiple conditions (e.g.
    /// <see cref="XLFilterType.Regular"/> can have many values while <see cref="XLFilterType.Dynamic"/>
    /// can be only one).
    /// </para>
    /// <para>
    /// <para>
    /// Whenever filter configuration changes, the filters are immediately reapplied.
    /// </para>
    /// <list type="bullet">
    ///   <item><term>Top/Bottom</term><description>only accept value in any of the highest/lowest values of the column.</description></item>
    ///   <item><term>Average</term><description>only accept value above/below average of all values in the column.</description></item>
    ///   <item><term>Text filters</term><description>only accept value whose text representation matches <see cref="Wildcard"/>. It encompasses text equality, <c>start-with</c> ect.</description></item>
    ///   <item><term>Number</term><description>only accept value whose text representation matches <see cref="Wildcard"/>.</description></item>
    /// </list>
    /// </para>
    /// </summary>
    public interface IXLFilterColumn
    {
        /// <summary>
        /// Remove all filters from the column.
        /// </summary>
        void Clear();

        /// <summary>
        /// <para>
        /// Switch to <see cref="XLFilterType.Regular"/> filter if necessary and add
        /// <paramref name="value"/> to a set of allowed values. Excel displays regular filter as
        /// a list of possible values in a column with checkbox next to it and user can check which
        /// one should be displayed.
        /// </para>
        /// <para>
        /// From technical perspective, the passed <paramref name="value"/> is converted to
        /// a localized string (using current locale) and the column values satisfy the filter
        /// condition, when its formatted string matches any filter string.
        /// </para>
        /// <para>
        /// Examples of less intuitive behavior: filter value is <c>2.5</c> in locale cs-CZ that
        /// uses "<em>,</em>" as a decimal separator. The passed <paramref name="value"/>
        /// is number 2.5, converted to string <em>2,5</em>. This string is used for comparison
        /// with values in the column:
        /// <list type="bullet">
        ///  <item>Number 2.5 formatted with two decimal places as <em>2,50</em> will not match.</item>
        ///  <item>Number 2.5 with default formatting will be matched, because its string is <em>2,5</em> in cs-CZ locale (but not in others, e.g. en-US locale).</item>
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

        IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping);

        void Top(Int32 value, XLTopBottomType type = XLTopBottomType.Items);

        void Bottom(Int32 value, XLTopBottomType type = XLTopBottomType.Items);

        void AboveAverage();

        void BelowAverage();

        IXLFilterConnector EqualTo<T>(T value) where T : IComparable<T>;

        IXLFilterConnector NotEqualTo<T>(T value) where T : IComparable<T>;

        IXLFilterConnector GreaterThan<T>(T value) where T : IComparable<T>;

        IXLFilterConnector LessThan<T>(T value) where T : IComparable<T>;

        IXLFilterConnector EqualOrGreaterThan<T>(T value) where T : IComparable<T>;

        IXLFilterConnector EqualOrLessThan<T>(T value) where T : IComparable<T>;

        void Between<T>(T minValue, T maxValue) where T : IComparable<T>;

        void NotBetween<T>(T minValue, T maxValue) where T : IComparable<T>;

        IXLFilterConnector BeginsWith(String value);

        IXLFilterConnector NotBeginsWith(String value);

        IXLFilterConnector EndsWith(String value);

        IXLFilterConnector NotEndsWith(String value);

        IXLFilterConnector Contains(String value);

        IXLFilterConnector NotContains(String value);

        XLFilterType FilterType { get; set; }
        Int32 TopBottomValue { get; set; }
        XLTopBottomType TopBottomType { get; set; }
        XLTopBottomPart TopBottomPart { get; set; }
        XLFilterDynamicType DynamicType { get; set; }
        Double DynamicValue { get; set; }

        IXLFilterColumn SetFilterType(XLFilterType value);

        IXLFilterColumn SetTopBottomValue(Int32 value);

        IXLFilterColumn SetTopBottomType(XLTopBottomType value);

        IXLFilterColumn SetTopBottomPart(XLTopBottomPart value);

        IXLFilterColumn SetDynamicType(XLFilterDynamicType value);

        IXLFilterColumn SetDynamicValue(Double value);
    }
}
