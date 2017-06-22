using System;

namespace ClosedXML.Excel
{
    public class XLWholeNumberCriteria : XLValidationCriteria
    {
        public XLWholeNumberCriteria(IXLDataValidation dataValidation): base(dataValidation)
        {
            
        }

        public void EqualTo(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualTo;
        }
        public void NotEqualTo(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.NotEqualTo;
        }
        public void GreaterThan(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.GreaterThan;
        }
        public void LessThan(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.LessThan;
        }
        public void EqualOrGreaterThan(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }
        public void EqualOrLessThan(Int32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }
        public void Between(Int32 minValue, Int32 maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.Between;
        }
        public void NotBetween(Int32 minValue, Int32 maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.NotBetween;
        }


    }
}
