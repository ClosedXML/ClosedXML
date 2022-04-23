// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLValidationCriteria
    {
        void Between(string minValue, string maxValue);

        [Obsolete("Use the overload accepting IXLCell")]
        void Between(IXLRange minValue, IXLRange maxValue);

        void Between(IXLCell minValue, IXLCell maxValue);

        void EqualOrGreaterThan(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void EqualOrGreaterThan(IXLRange range);

        void EqualOrGreaterThan(IXLCell cell);

        void EqualOrLessThan(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void EqualOrLessThan(IXLRange range);

        void EqualOrLessThan(IXLCell cell);

        void EqualTo(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void EqualTo(IXLRange range);

        void EqualTo(IXLCell cell);

        void GreaterThan(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void GreaterThan(IXLRange range);

        void GreaterThan(IXLCell cell);

        void LessThan(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void LessThan(IXLRange range);

        void LessThan(IXLCell cell);

        void NotBetween(string minValue, string maxValue);

        [Obsolete("Use the overload accepting IXLCell")]
        void NotBetween(IXLRange minValue, IXLRange maxValue);

        void NotBetween(IXLCell minValue, IXLCell maxValue);

        void NotEqualTo(string value);

        [Obsolete("Use the overload accepting IXLCell")]
        void NotEqualTo(IXLRange range);

        void NotEqualTo(IXLCell cell);
    }
}
