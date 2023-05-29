#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLValidationCriteria
    {
        void Between(String minValue, String maxValue);

        void Between(IXLCell minValue, IXLCell maxValue);

        void EqualOrGreaterThan(String value);

        void EqualOrGreaterThan(IXLCell cell);

        void EqualOrLessThan(String value);

        void EqualOrLessThan(IXLCell cell);

        void EqualTo(String value);

        void EqualTo(IXLCell cell);

        void GreaterThan(String value);

        void GreaterThan(IXLCell cell);

        void LessThan(String value);

        void LessThan(IXLCell cell);

        void NotBetween(String minValue, String maxValue);

        void NotBetween(IXLCell minValue, IXLCell maxValue);

        void NotEqualTo(String value);

        void NotEqualTo(IXLCell cell);
    }
}
