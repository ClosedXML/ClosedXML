// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class RangeEventArgs : EventArgs
    {
        public RangeEventArgs(IXLRange range)
        {
            Range = range ?? throw new ArgumentNullException(nameof(range));
        }

        public IXLRange Range { get; }
    }

    internal class XLDataValidation : IXLDataValidation
    {
        private readonly XLRanges _ranges;
        private readonly XLWorksheet _worksheet;

        public XLDataValidation(IXLRange range)
            : this(range?.Worksheet as XLWorksheet)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));

            AddRange(range);
        }

        public XLDataValidation(IXLDataValidation dataValidation, XLWorksheet worksheet)
            : this(worksheet)
        {
            _worksheet = worksheet;
            CopyFrom(dataValidation);
        }

        private XLDataValidation(XLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _ranges = new XLRanges();
            Initialize();
        }

        internal event EventHandler<RangeEventArgs> RangeAdded;

        internal event EventHandler<RangeEventArgs> RangeRemoved;

        internal XLWorksheet Worksheet => _worksheet;

        public void Clear()
        {
            Initialize();
        }

        public void CopyFrom(IXLDataValidation dataValidation)
        {
            if (dataValidation == this) return;

            if (!_ranges.Any())
                AddRanges(dataValidation.Ranges);

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

        public Boolean IsDirty()
        {
            return
                AllowedValues != XLAllowedValues.AnyValue
                || (ShowInputMessage &&
                   (!String.IsNullOrWhiteSpace(InputTitle) || !String.IsNullOrWhiteSpace(InputMessage)))
                || (ShowErrorMessage &&
                   (!String.IsNullOrWhiteSpace(ErrorTitle) || !String.IsNullOrWhiteSpace(ErrorMessage)));
        }

        internal void SplitBy(IXLRangeAddress rangeAddress)
        {
            var rangesToSplit = _ranges.GetIntersectedRanges(rangeAddress).ToList();

            foreach (var rangeToSplit in rangesToSplit)
            {
                var newRanges = (rangeToSplit as XLRange).Split(rangeAddress, includeIntersection: false);
                RemoveRange(rangeToSplit);
                newRanges.ForEach(AddRange);
            }
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

        #region IXLDataValidation Members

        private String maxValue;
        private String minValue;
        public XLAllowedValues AllowedValues { get; set; }

        public XLDateCriteria Date
        {
            get
            {
                AllowedValues = XLAllowedValues.Date;
                return new XLDateCriteria(this);
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

        public String ErrorMessage { get; set; }
        public XLErrorStyle ErrorStyle { get; set; }
        public String ErrorTitle { get; set; }
        public Boolean IgnoreBlanks { get; set; }
        public Boolean InCellDropdown { get; set; }
        public String InputMessage { get; set; }
        public String InputTitle { get; set; }
        public String MaxValue { get => maxValue; set { Validate(value); maxValue = value; } }
        public String MinValue { get => minValue; set { Validate(value); minValue = value; } }
        public XLOperator Operator { get; set; }
        public IEnumerable<IXLRange> Ranges => _ranges.AsEnumerable();

        public Boolean ShowErrorMessage { get; set; }

        public Boolean ShowInputMessage { get; set; }

        public XLTextLengthCriteria TextLength
        {
            get
            {
                AllowedValues = XLAllowedValues.TextLength;
                return new XLTextLengthCriteria(this);
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

        public String Value
        {
            get { return MinValue; }
            set { MinValue = value; }
        }

        public XLWholeNumberCriteria WholeNumber
        {
            get
            {
                AllowedValues = XLAllowedValues.WholeNumber;
                return new XLWholeNumberCriteria(this);
            }
        }

        /// <summary>
        /// Add a range to the collection of ranges this rule applies to.
        /// If the specified range does not belong to the worksheet of the data validation
        /// rule it is transferred to the target worksheet.
        /// </summary>
        /// <param name="range">A range to add.</param>
        public void AddRange(IXLRange range)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));

            if (range.Worksheet != Worksheet)
                range = Worksheet.Range(((XLRangeAddress)range.RangeAddress).WithoutWorksheet());

            _ranges.Add(range);

            RangeAdded?.Invoke(this, new RangeEventArgs(range));
        }

        /// <summary>
        /// Add a collection of ranges to the collection of ranges this rule applies to.
        /// Ranges that do not belong to the worksheet of the data validation
        /// rule are transferred to the target worksheet.
        /// </summary>
        /// <param name="ranges">Ranges to add.</param>
        public void AddRanges(IEnumerable<IXLRange> ranges)
        {
            ranges = ranges ?? Enumerable.Empty<IXLRange>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }
        }

        /// <summary>
        /// Detach data validation rule of all ranges it applies to.
        /// </summary>
        public void ClearRanges()
        {
            var allRanges = _ranges.ToList();
            _ranges.RemoveAll();

            foreach (var range in allRanges)
            {
                RangeRemoved?.Invoke(this, new RangeEventArgs(range));
            }
        }

        public void Custom(String customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
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

        /// <summary>
        /// Remove the specified range from the collection of range this rule applies to.
        /// </summary>
        /// <param name="range">A range to remove.</param>
        public bool RemoveRange(IXLRange range)
        {
            if (range == null)
                return false;

            var res = _ranges.Remove(range);

            if (res)
            {
                RangeRemoved?.Invoke(this, new RangeEventArgs(range));
            }

            return res;
        }

        #endregion IXLDataValidation Members

        private void Validate(String value)
        {
            if (value.Length > 255)
                throw new ArgumentOutOfRangeException(nameof(value), "The maximum allowed length of the value is 255 characters.");
        }
    }
}
