using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLDataValidations : IXLDataValidations, IEnumerable<XLDataValidation>, ISheetListener
    {
        private readonly List<XLDataValidation> _dataValidations = new();
        private readonly XLWorksheet _worksheet;

        public XLDataValidations(XLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        #region IXLDataValidations Members

        IXLWorksheet IXLDataValidations.Worksheet => _worksheet;

        IXLDataValidation IXLDataValidations.Add(IXLDataValidation dataValidation)
        {
            if (dataValidation == null)
                throw new ArgumentNullException(nameof(dataValidation));

            return Add((XLDataValidation)dataValidation);
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

        /// <summary>
        /// Get all data validation rules applied to ranges that intersect the specified range.
        /// </summary>
        public IEnumerable<IXLDataValidation> GetAllInRange(IXLRangeAddress rangeAddress)
        {
            if (rangeAddress == null || !rangeAddress.IsValid)
                yield break;

            var intersectingArea = XLSheetRange.FromRangeAddress(rangeAddress);

            foreach (var dataValidation in _dataValidations)
            {
                if (rangeAddress.Worksheet != dataValidation.Worksheet)
                    continue;

                foreach (var area in dataValidation.Areas)
                {
                    if (intersectingArea.Intersects(area))
                    {
                        yield return dataValidation;
                        break;
                    }
                }
            }
        }

        public IEnumerator<XLDataValidation> GetEnumerator()
        {
            return _dataValidations.GetEnumerator();
        }

        IEnumerator<IXLDataValidation> IEnumerable<IXLDataValidation>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Get the data validation rule for the range with the specified address if it exists.
        /// </summary>
        /// <param name="rangeAddress">A range address.</param>
        /// <param name="foundDataValidation">Data validation rule which ranges collection includes the specified
        /// address. The specified range should be fully covered with the data validation rule.
        /// For example, if the rule is applied to ranges A1:A3,C1:C3 then this method will
        /// return True for ranges A1:A3, C1:C2, A2:A3, and False for ranges A1:C3, A1:C1, etc.</param>
        /// <returns>True is the data validation rule was found, false otherwise.</returns>
        public bool TryGet(IXLRangeAddress rangeAddress, [NotNullWhen(true)] out IXLDataValidation? foundDataValidation)
        {
            if (rangeAddress == null || !rangeAddress.IsValid)
            {
                foundDataValidation = null;
                return false;
            }

            var coveredArea = XLSheetRange.FromRangeAddress(rangeAddress);

            foreach (var dataValidation in _dataValidations)
            {
                if (rangeAddress.Worksheet != dataValidation.Worksheet)
                    continue;

                foreach (var area in dataValidation.Areas)
                {
                    if (area.Covers(coveredArea))
                    {
                        foundDataValidation = dataValidation;
                        return true;
                    }
                }
            }

            foundDataValidation = null;
            return false;
        }

        #endregion IXLDataValidations Members

        internal XLDataValidation Add(XLDataValidation dataValidation)
        {
            return Add(dataValidation, skipIntersectionsCheck: false);
        }

        internal XLDataValidation Add(XLDataValidation dataValidation, bool skipIntersectionsCheck)
        {
            XLDataValidation xlDataValidation;
            if (dataValidation.Ranges.Any(r => r.Worksheet != _worksheet))
            {
                xlDataValidation = new XLDataValidation(dataValidation, _worksheet);
            }
            else
            {
                xlDataValidation = dataValidation;
            }

            // Adding a range can split current one
            foreach (var area in dataValidation.Areas)
                AdjustDataValidationAreas(_worksheet, area, static (dataValidationAreas, areaOfNewValidation) => dataValidationAreas.DeleteWithoutShift(areaOfNewValidation));

            _dataValidations.Add(xlDataValidation);
            return xlDataValidation;
        }

        internal void Delete(XLBookArea bookArea)
        {
            for (var i = _dataValidations.Count - 1; i >= 0; --i)
            {
                var dataValidation = _dataValidations[i];
                if (!XLHelper.SheetComparer.Equals(dataValidation.Worksheet.Name, bookArea.Name))
                    continue;

                foreach (var area in dataValidation.Areas)
                {
                    if (area.Intersects(bookArea.Area))
                    {
                        _dataValidations.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        internal void Delete(XLDataValidation dataValidation)
        {
            _dataValidations.Remove(dataValidation);
        }

        internal void Consolidate()
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

                IXLRanges consolidatedRanges = new XLRanges(_worksheet);
                ranges.ForEach(r => consolidatedRanges.Add(r));
                consolidatedRanges = consolidatedRanges.Consolidate();

                consRule.ClearRanges();
                consRule.AddRanges(consolidatedRanges);
                Add(consRule);
            }
        }

        void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, XLSheetRange area)
        {
            AdjustDataValidationAreas(sheet, area, static (sqref, insertedArea) => sqref.InsertAndShiftDown(insertedArea));
        }

        void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, XLSheetRange area)
        {
            AdjustDataValidationAreas(sheet, area, static (sqref, insertedArea) => sqref.InsertAndShiftRight(insertedArea));
        }

        void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, XLSheetRange deletedRange)
        {
            AdjustDataValidationAreas(sheet, deletedRange, static (sqref, deletedArea) => sqref.DeleteAndShiftLeft(deletedArea));
        }

        void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, XLSheetRange deletedRange)
        {
            AdjustDataValidationAreas(sheet, deletedRange, static (sqref, deletedArea) => sqref.DeleteAndShiftUp(deletedArea));
        }

        private void AdjustDataValidationAreas(XLWorksheet sheet, XLSheetRange affectedRange, Func<XLAreaList, XLSheetRange, XLAreaList> adjustAreas)
        {
            if (sheet != _worksheet)
                return;

            for (var i = _dataValidations.Count - 1; i >= 0; --i)
            {
                var dataValidation = _dataValidations[i];
                var modifiedAreaList = adjustAreas(dataValidation.Areas, affectedRange);
                if (modifiedAreaList.Count == 0)
                {
                    _dataValidations.RemoveAt(i);
                }
                else
                {
                    dataValidation.Areas = modifiedAreaList;
                }
            }
        }
    }
}
