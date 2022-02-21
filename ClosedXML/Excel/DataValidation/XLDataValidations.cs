// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel.Ranges.Index;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    using System.Collections;
    using System.Linq;

    internal class XLDataValidations : IXLDataValidations
    {
        private readonly XLRangeIndex<XLDataValidationIndexEntry> _dataValidationIndex;

        private readonly List<IXLDataValidation> _dataValidations = new List<IXLDataValidation>();
        private readonly XLWorksheet _worksheet;

        /// <summary>
        /// The flag used to avoid unnecessary check for splitting intersected ranges when we already
        /// are performing the splitting.
        /// </summary>
        private bool _skipSplittingExistingRanges = false;

        public XLDataValidations(XLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _dataValidationIndex = new XLRangeIndex<XLDataValidationIndexEntry>(_worksheet);
        }

        internal XLWorksheet Worksheet => _worksheet;

        #region IXLDataValidations Members

        IXLWorksheet IXLDataValidations.Worksheet => _worksheet;

        public IXLDataValidation Add(IXLDataValidation dataValidation)
        {
            return Add(dataValidation, skipIntersectionsCheck: false);
        }

        public Boolean ContainsSingle(IXLRange range)
        {
            Int32 count = 0;
            foreach (var xlDataValidation in _dataValidations.Where(dv => dv.Ranges.Contains(range)))
            {
                count++;
                if (count > 1) return false;
            }

            return count == 1;
        }

        public void Delete(Predicate<IXLDataValidation> predicate)
        {
            var dataValidationsToRemove = _dataValidations.Where(dv => predicate(dv))
                .ToList();

            dataValidationsToRemove.ForEach(Delete);
        }

        public void Delete(IXLDataValidation dataValidation)
        {
            if (!_dataValidations.Remove(dataValidation))
                return;
            var xlDataValidation = dataValidation as XLDataValidation;
            xlDataValidation.RangeAdded -= OnRangeAdded;
            xlDataValidation.RangeRemoved -= OnRangeRemoved;

            foreach (var range in dataValidation.Ranges)
            {
                ProcessRangeRemoved(range);
            }
        }

        public void Delete(IXLRange range)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));

            var dataValidationsToRemove = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)range.RangeAddress)
                .Select(e => e.DataValidation)
                .Distinct()
                .ToList();

            dataValidationsToRemove.ForEach(Delete);
        }

        /// <summary>
        /// Get all data validation rules applied to ranges that intersect the specified range.
        /// </summary>
        public IEnumerable<IXLDataValidation> GetAllInRange(IXLRangeAddress rangeAddress)
        {
            if (rangeAddress == null || !rangeAddress.IsValid)
                return Enumerable.Empty<IXLDataValidation>();

            return _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
                .Select(indexEntry => indexEntry.DataValidation)
                .Distinct();
        }

        public IEnumerator<IXLDataValidation> GetEnumerator()
        {
            return _dataValidations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Get the data validation rule for the range with the specified address if it exists.
        /// </summary>
        /// <param name="rangeAddress">A range address.</param>
        /// <param name="dataValidation">Data validation rule which ranges collection includes the specified
        /// address. The specified range should be fully covered with the data validation rule.
        /// For example, if the rule is applied to ranges A1:A3,C1:C3 then this method will
        /// return True for ranges A1:A3, C1:C2, A2:A3, and False for ranges A1:C3, A1:C1, etc.</param>
        /// <returns>True is the data validation rule was found, false otherwise.</returns>
        public bool TryGet(IXLRangeAddress rangeAddress, out IXLDataValidation dataValidation)
        {
            dataValidation = null;
            if (rangeAddress == null || !rangeAddress.IsValid)
                return false;

            var candidates = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
                .Where(c => c.RangeAddress.Contains(rangeAddress.FirstAddress) &&
                            c.RangeAddress.Contains(rangeAddress.LastAddress));

            if (!candidates.Any())
                return false;

            dataValidation = candidates.First().DataValidation;

            return true;
        }

        internal IXLDataValidation Add(IXLDataValidation dataValidation, bool skipIntersectionsCheck)
        {
            if (dataValidation == null) throw new ArgumentNullException(nameof(dataValidation));

            XLDataValidation xlDataValidation;
            if (!(dataValidation is XLDataValidation) ||
                dataValidation.Ranges.Any(r => r.Worksheet != Worksheet))
            {
                xlDataValidation = new XLDataValidation(dataValidation, Worksheet);
            }
            else
            {
                xlDataValidation = (XLDataValidation)dataValidation;
            }

            xlDataValidation.RangeAdded += OnRangeAdded;
            xlDataValidation.RangeRemoved += OnRangeRemoved;

            foreach (var range in xlDataValidation.Ranges)
            {
                ProcessRangeAdded(range, xlDataValidation, skipIntersectionsCheck);
            }

            _dataValidations.Add(xlDataValidation);

            return xlDataValidation;
        }

        #endregion IXLDataValidations Members

        public void Consolidate()
        {
            Func<IXLDataValidation, IXLDataValidation, bool> areEqual = (dv1, dv2) =>
            {
                return
                    dv1.IgnoreBlanks == dv2.IgnoreBlanks &&
                    dv1.InCellDropdown == dv2.InCellDropdown &&
                    dv1.ShowErrorMessage == dv2.ShowErrorMessage &&
                    dv1.ShowInputMessage == dv2.ShowInputMessage &&
                    dv1.InputTitle == dv2.InputTitle &&
                    dv1.InputMessage == dv2.InputMessage &&
                    dv1.ErrorTitle == dv2.ErrorTitle &&
                    dv1.ErrorMessage == dv2.ErrorMessage &&
                    dv1.ErrorStyle == dv2.ErrorStyle &&
                    dv1.AllowedValues == dv2.AllowedValues &&
                    dv1.Operator == dv2.Operator &&
                    dv1.MinValue == dv2.MinValue &&
                    dv1.MaxValue == dv2.MaxValue &&
                    dv1.Value == dv2.Value;
            };

            var rules = _dataValidations.ToList();
            rules.ForEach(Delete);

            while (rules.Any())
            {
                var similarRules = rules.Where(r => areEqual(rules.First(), r)).ToList();
                similarRules.ForEach(r => rules.Remove(r));

                var consRule = similarRules.First();
                var ranges = similarRules.SelectMany(dv => dv.Ranges).ToList();

                IXLRanges consolidatedRanges = new XLRanges();
                ranges.ForEach(r => consolidatedRanges.Add(r));
                consolidatedRanges = consolidatedRanges.Consolidate();

                consRule.ClearRanges();
                consRule.AddRanges(consolidatedRanges);
                Add(consRule);
            }
        }

        private void OnRangeAdded(object sender, RangeEventArgs e)
        {
            ProcessRangeAdded(e.Range, sender as XLDataValidation, skipIntersectionCheck: false);
        }

        private void OnRangeRemoved(object sender, RangeEventArgs e)
        {
            ProcessRangeRemoved(e.Range);
        }

        private void ProcessRangeAdded(IXLRange range, XLDataValidation dataValidation, bool skipIntersectionCheck)
        {
            if (!skipIntersectionCheck)
            {
                SplitExistingRanges(range.RangeAddress);
            }

            var indexEntry = new XLDataValidationIndexEntry(range.RangeAddress, dataValidation);
            _dataValidationIndex.Add(indexEntry);
        }

        private void ProcessRangeRemoved(IXLRange range)
        {
            var entry = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)range.RangeAddress)
                .SingleOrDefault(e => Equals(e.RangeAddress, range.RangeAddress));

            if (entry != null)
            {
                _dataValidationIndex.Remove(entry.RangeAddress);
            }
        }

        private void SplitExistingRanges(IXLRangeAddress rangeAddress)
        {
            if (_skipSplittingExistingRanges) return;

            try
            {
                _skipSplittingExistingRanges = true;
                var entries = _dataValidationIndex.GetIntersectedRanges((XLRangeAddress)rangeAddress)
                    .ToList();

                foreach (var entry in entries)
                {
                    entry.DataValidation.SplitBy(rangeAddress);
                }
            }
            finally
            {
                _skipSplittingExistingRanges = false;
            }

            //TODO Remove empty data validations
        }

        /// <summary>
        /// Class used for indexing data validation rules.
        /// </summary>
        private class XLDataValidationIndexEntry : IXLAddressable
        {
            public XLDataValidationIndexEntry(IXLRangeAddress rangeAddress, XLDataValidation dataValidation)
            {
                RangeAddress = rangeAddress;
                DataValidation = dataValidation;
            }

            public XLDataValidation DataValidation { get; }

            /// <summary>
            ///   Gets an object with the boundaries of this range.
            /// </summary>
            public IXLRangeAddress RangeAddress { get; }
        }
    }
}
