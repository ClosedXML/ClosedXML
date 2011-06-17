using System;

namespace ClosedXML.Excel
{
    internal class XLDataValidation: IXLDataValidation
    {
        private XLWorksheet worksheet;
        public XLDataValidation(IXLRanges ranges, XLWorksheet worksheet)
        {
            this.Ranges = ranges;
            this.AllowedValues = XLAllowedValues.AnyValue;
            this.IgnoreBlanks = true;
            ShowErrorMessage = true;
            ShowInputMessage = true;
            InCellDropdown = true;
            Operator = XLOperator.Between;
            this.worksheet = worksheet;
        }

        public XLDataValidation(IXLDataValidation dataValidation, XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
            this.CopyFrom(dataValidation);
        }

        public IXLRanges Ranges { get; set; }
        
        public void Delete()
        {
            throw new NotImplementedException();
            //foreach (var dv in worksheet.DataValidations)
            //{ 

            //}
        }
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
            Value = list;
        }
        public void List(IXLRange range)
        {
            List(range, true);
        }
        public void List(IXLRange range, Boolean inCellDropdown)
        {
            AllowedValues = XLAllowedValues.List;
            Value = String.Format("'{0}'!{1}", ((XLRange)range).Worksheet.Name, range.RangeAddress.ToString());
        }

        public void Custom(String customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
        }
    }
}
