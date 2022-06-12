// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.Ranges;
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
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

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
            if (dataValidation == this)
            {
                return;
            }

            if (!_ranges.Any())
            {
                AddRanges(dataValidation.Ranges);
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

        public bool IsDirty()
        {
            return
                AllowedValues != XLAllowedValues.AnyValue
                || (ShowInputMessage &&
                   (!string.IsNullOrWhiteSpace(InputTitle) || !string.IsNullOrWhiteSpace(InputMessage)))
                || (ShowErrorMessage &&
                   (!string.IsNullOrWhiteSpace(ErrorTitle) || !string.IsNullOrWhiteSpace(ErrorMessage)));
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
            InputTitle = string.Empty;
            InputMessage = string.Empty;
            ErrorTitle = string.Empty;
            ErrorMessage = string.Empty;
            ErrorStyle = XLErrorStyle.Stop;
            Operator = XLOperator.Between;
            Value = string.Empty;
            MinValue = string.Empty;
            MaxValue = string.Empty;
        }

        #region IXLDataValidation Members

        private string maxValue;
        private string minValue;
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

        public string ErrorMessage { get; set; }
        public XLErrorStyle ErrorStyle { get; set; }
        public string ErrorTitle { get; set; }
        public bool IgnoreBlanks { get; set; }
        public bool InCellDropdown { get; set; }
        public string InputMessage { get; set; }
        public string InputTitle { get; set; }
        public string MaxValue { get => maxValue; set { Validate(value); maxValue = value; } }
        public string MinValue { get => minValue; set { Validate(value); minValue = value; } }
        public XLOperator Operator { get; set; }
        public IEnumerable<IXLRange> Ranges => _ranges.AsEnumerable();

        public bool ShowErrorMessage { get; set; }

        public bool ShowInputMessage { get; set; }

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

        public string Value
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
            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            if (range.Worksheet != Worksheet)
            {
                range = Worksheet.Range(((XLRangeAddress)range.RangeAddress).WithoutWorksheet());
            }

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

        public void Custom(string customValidation)
        {
            AllowedValues = XLAllowedValues.Custom;
            Value = customValidation;
        }

        public void List(string list)
        {
            List(list, true);
        }

        public void List(string list, bool inCellDropdown)
        {
            AllowedValues = XLAllowedValues.List;
            InCellDropdown = inCellDropdown;
            Value = list;
        }

        public void List(IXLRange range)
        {
            List(range, true);
        }

        public void List(IXLRange range, bool inCellDropdown)
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
            {
                return false;
            }

            var res = _ranges.Remove(range);

            if (res)
            {
                RangeRemoved?.Invoke(this, new RangeEventArgs(range));
            }

            return res;
        }

        #endregion IXLDataValidation Members

        private void Validate(string value)
        {
            if (value.Length > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The maximum allowed length of the value is 255 characters.");
            }
        }
    }
}
