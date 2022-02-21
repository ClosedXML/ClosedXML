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

        public void Between(TimeSpan minValue, TimeSpan maxValue) => base.Between(GetXLTime(minValue), GetXLTime(maxValue));

        public void EqualOrGreaterThan(TimeSpan value) => base.EqualOrGreaterThan(GetXLTime(value));

        public void EqualOrLessThan(TimeSpan value) => base.EqualOrLessThan(GetXLTime(value));

        public void EqualTo(TimeSpan value) => base.EqualTo(GetXLTime(value));

        public void GreaterThan(TimeSpan value) => base.GreaterThan(GetXLTime(value));

        public void LessThan(TimeSpan value) => base.LessThan(GetXLTime(value));

        public void NotBetween(TimeSpan minValue, TimeSpan maxValue) => base.NotBetween(GetXLTime(minValue), GetXLTime(maxValue));

        public void NotEqualTo(TimeSpan value) => base.NotEqualTo(GetXLTime(value));

        private static String GetXLTime(TimeSpan value)
        {
            return (value.TotalHours / 24.0).ToString();
        }
    }
}
