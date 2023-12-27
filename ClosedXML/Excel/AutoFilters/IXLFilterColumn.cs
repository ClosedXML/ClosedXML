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
        void Clear();

        IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>;

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
