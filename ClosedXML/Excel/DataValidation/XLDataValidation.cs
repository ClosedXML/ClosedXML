using System;

namespace ClosedXML.Excel
{
    internal class XLDataValidation : IXLDataValidation
    {
        public XLDataValidation(IXLRanges ranges)
        {
            
            Ranges = new XLRanges();
            ranges.ForEach(r=>
                               {
                                   var newR =
                                       new XLRange(new XLRangeParameters(r.RangeAddress as XLRangeAddress,
                                                                         r.Worksheet.Style) {IgnoreEvents = true});
                                   (Ranges as XLRanges).Add(newR);
                               } );
            Initialize();
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
                  (!XLHelper.IsNullOrWhiteSpace(InputTitle) || !XLHelper.IsNullOrWhiteSpace(InputMessage)))
                ||(ShowErrorMessage &&
                  (!XLHelper.IsNullOrWhiteSpace(ErrorTitle) || !XLHelper.IsNullOrWhiteSpace(ErrorMessage)));

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
        private XLAllowedValues _allowedValues;
        public XLAllowedValues AllowedValues
        {
            get { return _allowedValues; }
            set { _allowedValues = value; }
        }
        
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
            List(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
        }

        public void Custom(String customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
        }

        #endregion

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
    }
}