using System;

namespace ClosedXML.Excel
{
    internal class XLDataValidation : IXLDataValidation
    {
        private XLDataValidation()
        {
            Ranges = new XLRanges();
            Initialize();
        }

        public XLDataValidation(IXLRange range)
            : this()
        {
            Ranges.Add(new XLRange(new XLRangeParameters((XLRangeAddress)range.RangeAddress, range.Worksheet.Style)));
        }

        public XLDataValidation(IXLRanges ranges)
            : this()
        {
            ranges.ForEach(range =>
            {
                Ranges.Add(new XLRange(new XLRangeParameters((XLRangeAddress)range.RangeAddress, range.Worksheet.Style)));
            });
        }

        private void Initialize()
        {
            AllowedValues = XLAllowedValues.AnyValue;
            IgnoreBlanks = true;
            ShowErrorMessage = true;
            ShowInputMessage = true;
            InCellDropdown = true;
            InputTitle = String.Empty;
            InputMessage = String.Empty;
            ErrorTitle = String.Empty;
            ErrorMessage = String.Empty;
            ErrorStyle = XLErrorStyle.Stop;
            Operator = XLOperator.Between;
            Value = String.Empty;
            MinValue = String.Empty;
            MaxValue = String.Empty;
        }

        public Boolean IsDirty()
        {
            return
                AllowedValues != XLAllowedValues.AnyValue
                || (ShowInputMessage &&
                   (!String.IsNullOrWhiteSpace(InputTitle) || !String.IsNullOrWhiteSpace(InputMessage)))
                || (ShowErrorMessage &&
                   (!String.IsNullOrWhiteSpace(ErrorTitle) || !String.IsNullOrWhiteSpace(ErrorMessage)));
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

        private String minValue;
        public String MinValue { get => minValue; set { Validate(value); minValue = value; } }

        private String maxValue;
        public String MaxValue { get => maxValue; set { Validate(value); maxValue = value; } }

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
            List(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
        }

        public void Custom(String customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
        }

        #endregion IXLDataValidation Members

        public void CopyFrom(IXLDataValidation dataValidation)
        {
            if (dataValidation == this) return;

            if (Ranges == null && dataValidation.Ranges != null)
            {
                Ranges = new XLRanges();
                dataValidation.Ranges.ForEach(r => Ranges.Add(r));
            }

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

        public void Clear()
        {
            Initialize();
        }

        private void Validate(String value)
        {
            if (value.Length > 255)
                throw new ArgumentOutOfRangeException(nameof(value), "The maximum allowed length of the value is 255 characters.");
        }
    }
}
