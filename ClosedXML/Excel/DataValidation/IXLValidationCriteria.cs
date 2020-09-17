using System;

namespace ClosedXML.Excel
{
    public interface IXLValidationCriteria
    {
        void EqualTo(String value);
        void NotEqualTo(String value);
        void GreaterThan(String value);
        void LessThan(String value);
        void EqualOrGreaterThan(String value);
        void EqualOrLessThan(String value);
        void Between(String minValue, String maxValue);
        void NotBetween(String minValue, String maxValue);

        [Obsolete("Use the overload accepting IXLCell")]
        void EqualTo(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void NotEqualTo(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void GreaterThan(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void LessThan(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void EqualOrGreaterThan(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void EqualOrLessThan(IXLRange range);
        [Obsolete("Use the overload accepting IXLCell")]
        void Between(IXLRange minValue, IXLRange maxValue);
        [Obsolete("Use the overload accepting IXLCell")]
        void NotBetween(IXLRange minValue, IXLRange maxValue);

        void EqualTo(IXLCell cell);
        void NotEqualTo(IXLCell cell);
        void GreaterThan(IXLCell cell);
        void LessThan(IXLCell cell);
        void EqualOrGreaterThan(IXLCell cell);
        void EqualOrLessThan(IXLCell cell);
        void Between(IXLCell minValue, IXLCell maxValue);
        void NotBetween(IXLCell minValue, IXLCell maxValue);
    }
}
