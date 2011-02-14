using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLTextLengthCriteria : XLValidationCriteria
    {
        public XLTextLengthCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
            
        }

        public void EqualTo(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualTo;
        }
        public void NotEqualTo(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.NotEqualTo;
        }
        public void GreaterThan(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.GreaterThan;
        }
        public void LessThan(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.LessThan;
        }
        public void EqualOrGreaterThan(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }
        public void EqualOrLessThan(UInt32 value)
        {
            dataValidation.Value = value.ToString();
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }
        public void Between(UInt32 minValue, UInt32 maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.Between;
        }
        public void NotBetween(UInt32 minValue, UInt32 maxValue)
        {
            dataValidation.MinValue = minValue.ToString();
            dataValidation.MaxValue = maxValue.ToString();
            dataValidation.Operator = XLOperator.NotBetween;
        }


    }
}
