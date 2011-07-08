using System;

namespace ClosedXML.Excel
{
    internal class XLDataValidation : IXLDataValidation
    {
        public XLDataValidation(IXLRanges ranges)
        {
            Ranges = ranges;
            AllowedValues = XLAllowedValues.AnyValue;
            IgnoreBlanks = true;
            ShowErrorMessage = true;
            ShowInputMessage = true;
            InCellDropdown = true;
            Operator = XLOperator.Between;
        }

        public XLDataValidation(IXLDataValidation dataValidation)
        {
            CopyFrom(dataValidation);
        }

        #region IXLDataValidation Members

        public IXLRanges Ranges { get; set; }


        public Boolean IgnoreBlanks { get; set; }
        public Boolean InCellDropdown { get; set; }
        public Boolean ShowInputMessage { get; set; }
        public String InputTitle { get; set; }
        public String InputMessage { get; set; }
        public Boolean ShowErrorMessage { get; set; }
        public String ErrorTitle { get; set; }
        public String ErrorMessage { get; set; }
        public XLErrorStyle ErrorStyle { get; set; }
        public XLAllowedValues AllowedValues { get; set; }
        public XLOperator Operator { get; set; }

        public String Value
        {
            get { return MinValue; }
            set { MinValue = value; }
        }

        public String MinValue { get; set; }
        public String MaxValue { get; set; }

        public XLWholeNumberCriteria WholeNumber
        {
            get
            {
                AllowedValues = XLAllowedValues.WholeNumber;
                return new XLWholeNumberCriteria(this);
            }
        }

        public XLDecimalCriteria Decimal
        {
            get
            {
                AllowedValues = XLAllowedValues.Decimal;
                return new XLDecimalCriteria(this);
            }
        }

        public XLDateCriteria Date
        {
            get
            {
                AllowedValues = XLAllowedValues.Date;
                return new XLDateCriteria(this);
            }
        }

        public XLTimeCriteria Time
        {
            get
            {
                AllowedValues = XLAllowedValues.Time;
                return new XLTimeCriteria(this);
            }
        }

        public XLTextLengthCriteria TextLength
        {
            get
            {
                AllowedValues = XLAllowedValues.TextLength;
                return new XLTextLengthCriteria(this);
            }
        }

        public void List(String list)
        {
            List(list, true);
        }

        public void List(String list, Boolean inCellDropdown)
        {
            AllowedValues = XLAllowedValues.List;
            InCellDropdown = inCellDropdown;
            Value = list;
        }

        public void List(IXLRange range)
        {
            List(range, true);
        }

        public void List(IXLRange range, Boolean inCellDropdown)
        {
            AllowedValues = XLAllowedValues.List;
            InCellDropdown = inCellDropdown;
            Value = String.Format("'{0}'!{1}", ((XLRange)range).Worksheet.Name, range.RangeAddress.ToStringFixed());
        }

        public void Custom(String customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
        }

        #endregion

        public void CopyFrom(IXLDataValidation dataValidation)
        {
            if (Ranges == null)
                Ranges = new XLRanges();
            //dataValidation.Ranges.ForEach(r => Ranges.Add(r));

            IgnoreBlanks = dataValidation.IgnoreBlanks;
            InCellDropdown = dataValidation.InCellDropdown;
            ShowErrorMessage = dataValidation.ShowErrorMessage;
            ShowInputMessage = dataValidation.ShowInputMessage;
            InputTitle = dataValidation.InputTitle;
            InputMessage = dataValidation.InputMessage;
            ErrorTitle = dataValidation.ErrorTitle;
            ErrorMessage = dataValidation.ErrorMessage;
            ErrorStyle = dataValidation.ErrorStyle;
            AllowedValues = dataValidation.AllowedValues;
            Operator = dataValidation.Operator;
            MinValue = dataValidation.MinValue;
            MaxValue = dataValidation.MaxValue;
        }
    }
}