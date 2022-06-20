using System;

namespace ClosedXML.Excel
{
    public enum XLTopBottomType { Items, Percent }

    public enum XLDateTimeGrouping { Year, Month, Day, Hour, Minute, Second }

    public interface IXLFilterColumn
    {
        void Clear();

        IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>;

        IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping);

        void Top(int value, XLTopBottomType type = XLTopBottomType.Items);

        void Bottom(int value, XLTopBottomType type = XLTopBottomType.Items);

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

        IXLFilterConnector BeginsWith(string value);

        IXLFilterConnector NotBeginsWith(string value);

        IXLFilterConnector EndsWith(string value);

        IXLFilterConnector NotEndsWith(string value);

        IXLFilterConnector Contains(string value);

        IXLFilterConnector NotContains(string value);

        XLFilterType FilterType { get; set; }
        int TopBottomValue { get; set; }
        XLTopBottomType TopBottomType { get; set; }
        XLTopBottomPart TopBottomPart { get; set; }
        XLFilterDynamicType DynamicType { get; set; }
        double DynamicValue { get; set; }

        IXLFilterColumn SetFilterType(XLFilterType value);

        IXLFilterColumn SetTopBottomValue(int value);

        IXLFilterColumn SetTopBottomType(XLTopBottomType value);

        IXLFilterColumn SetTopBottomPart(XLTopBottomPart value);

        IXLFilterColumn SetDynamicType(XLFilterDynamicType value);

        IXLFilterColumn SetDynamicValue(double value);
    }
}
