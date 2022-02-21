// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public abstract class XLValidationCriteria : IXLValidationCriteria
    {
        protected IXLDataValidation dataValidation;

        protected XLValidationCriteria(IXLDataValidation dataValidation)
        {
            this.dataValidation = dataValidation;
        }

        #region IXLValidationCriteria Members

        public void Between(String minValue, String maxValue)
        {
            dataValidation.MinValue = minValue;
            dataValidation.MaxValue = maxValue;
            dataValidation.Operator = XLOperator.Between;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void Between(IXLRange minValue, IXLRange maxValue)
        {
            dataValidation.MinValue = minValue.RangeAddress.ToStringFixed();
            dataValidation.MaxValue = maxValue.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.Between;
        }

        public void Between(IXLCell minValue, IXLCell maxValue)
        {
            dataValidation.MinValue = minValue.Address.ToStringFixed();
            dataValidation.MaxValue = maxValue.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.Between;
        }

        public void EqualOrGreaterThan(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void EqualOrGreaterThan(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        public void EqualOrGreaterThan(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualOrGreaterThan;
        }

        public void EqualOrLessThan(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void EqualOrLessThan(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        public void EqualOrLessThan(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualOrLessThan;
        }

        public void EqualTo(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.EqualTo;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void EqualTo(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualTo;
        }

        public void EqualTo(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.EqualTo;
        }

        public void GreaterThan(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void GreaterThan(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        public void GreaterThan(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.GreaterThan;
        }

        public void LessThan(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.LessThan;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void LessThan(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.LessThan;
        }

        public void LessThan(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.LessThan;
        }

        public void NotBetween(String minValue, String maxValue)
        {
            dataValidation.MinValue = minValue;
            dataValidation.MaxValue = maxValue;
            dataValidation.Operator = XLOperator.NotBetween;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void NotBetween(IXLRange minValue, IXLRange maxValue)
        {
            dataValidation.MinValue = minValue.RangeAddress.ToStringFixed();
            dataValidation.MaxValue = maxValue.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.NotBetween;
        }

        public void NotBetween(IXLCell minValue, IXLCell maxValue)
        {
            dataValidation.MinValue = minValue.Address.ToStringFixed();
            dataValidation.MaxValue = maxValue.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.NotBetween;
        }

        public void NotEqualTo(String value)
        {
            dataValidation.Value = value;
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        [Obsolete("Use the overload accepting IXLCell")]
        public void NotEqualTo(IXLRange range)
        {
            dataValidation.Value = range.RangeAddress.ToStringFixed();
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        public void NotEqualTo(IXLCell cell)
        {
            dataValidation.Value = cell.Address.ToStringFixed();
            dataValidation.Operator = XLOperator.NotEqualTo;
        }

        #endregion IXLValidationCriteria Members
    }
}
