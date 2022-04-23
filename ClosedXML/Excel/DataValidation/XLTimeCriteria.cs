// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public class XLTimeCriteria : XLValidationCriteria
    {
        public XLTimeCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void Between(TimeSpan minValue, TimeSpan maxValue) => Between(GetXLTime(minValue), GetXLTime(maxValue));

        public void EqualOrGreaterThan(TimeSpan value) => EqualOrGreaterThan(GetXLTime(value));

        public void EqualOrLessThan(TimeSpan value) => EqualOrLessThan(GetXLTime(value));

        public void EqualTo(TimeSpan value) => EqualTo(GetXLTime(value));

        public void GreaterThan(TimeSpan value) => GreaterThan(GetXLTime(value));

        public void LessThan(TimeSpan value) => LessThan(GetXLTime(value));

        public void NotBetween(TimeSpan minValue, TimeSpan maxValue) => NotBetween(GetXLTime(minValue), GetXLTime(maxValue));

        public void NotEqualTo(TimeSpan value) => NotEqualTo(GetXLTime(value));

        private static string GetXLTime(TimeSpan value)
        {
            return (value.TotalHours / 24.0).ToString();
        }
    }
}
