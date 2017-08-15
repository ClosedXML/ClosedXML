using System;

namespace ClosedXML.Excel
{
    public class XLDecimalCriteria : XLValidationCriteria
    {
        public XLDecimalCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void EqualTo(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualTo;
        }

        public void NotEqualTo(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        public void GreaterThan(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        public void LessThan(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.LessThan;
        }

        public void EqualOrGreaterThan(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        public void EqualOrLessThan(Double value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        public void Between(Double minValue, Double maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.Between;
        }

        public void NotBetween(Double minValue, Double maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.NotBetween;
        }
    }
}