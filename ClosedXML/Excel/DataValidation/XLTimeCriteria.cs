using System;

namespace ClosedXML.Excel
{
    public class XLTimeCriteria : XLValidationCriteria
    {
        public XLTimeCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        private static String GetXLTime(TimeSpan value)
        {
            return (value.TotalHours / 24.0).ToString();
        }

        public void EqualTo(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.EqualTo;
        }

        public void NotEqualTo(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        public void GreaterThan(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        public void LessThan(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.LessThan;
        }

        public void EqualOrGreaterThan(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        public void EqualOrLessThan(TimeSpan value)
        {
            dataValidation.Value = GetXLTime(value);
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        public void Between(TimeSpan minValue, TimeSpan maxValue)
        {
            dataValidation.MinValue = GetXLTime(minValue);
            dataValidation.MaxValue = GetXLTime(maxValue);
            dataValidation.Operator = XLOperator.Between;
        }

        public void NotBetween(TimeSpan minValue, TimeSpan maxValue)
        {
            dataValidation.MinValue = GetXLTime(minValue);
            dataValidation.MaxValue = GetXLTime(maxValue);
            dataValidation.Operator = XLOperator.NotBetween;
        }
    }
}