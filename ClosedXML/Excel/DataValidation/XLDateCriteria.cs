using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    public class XLDateCriteria : XLValidationCriteria
    {
        public XLDateCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void EqualTo(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.EqualTo;
        }

        public void NotEqualTo(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        public void GreaterThan(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        public void LessThan(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.LessThan;
        }

        public void EqualOrGreaterThan(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        public void EqualOrLessThan(DateTime value)
        {
            dataValidation.Value = value.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        public void Between(DateTime minValue, DateTime maxValue)
        {
            dataValidation.MinValue = minValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.MaxValue = maxValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.Between;
        }

        public void NotBetween(DateTime minValue, DateTime maxValue)
        {
            dataValidation.MinValue = minValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.MaxValue = maxValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            dataValidation.Operator = XLOperator.NotBetween;
        }
    }
}