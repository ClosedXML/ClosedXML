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

            var dv = (XLDataValidation)dataValidation;
            if (dv.Worksheet != _worksheet)
                return CopyFrom(dv);

            // It's possible that it was detached and while detached, it had added some areas?
            // I have a very hard time understanding the use case and intended behavior. This
            // API should be scrapped.
            if (!_dataValidations.Contains(dv))
            {
                // Adding a range can split current one -> clear existing DVs so new one can be
                // added and "one DV per cell" is kept.
                foreach (var area in dv.Areas)
                    AdjustDataValidationAreas(_worksheet, area, static (dataValidationAreas, areaOfNewValidation) => dataValidationAreas.DeleteWithoutShift(areaOfNewValidation));

                _dataValidations.Add(dv);
            }

            return dv;
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
            if (rangeAddress is null || !rangeAddress.IsValid)
                yield break;

            if (rangeAddress.Worksheet != _worksheet)
                yield break;

            var intersectingArea = XLSheetRange.FromRangeAddress(rangeAddress);
            foreach (var dataValidation in _dataValidations)
            {
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
            if (rangeAddress is null || !rangeAddress.IsValid || rangeAddress.Worksheet != _worksheet)
            {
                foundDataValidation = null;
                return false;
            }

            var coveredArea = XLSheetRange.FromRangeAddress(rangeAddress);
            foreach (var dataValidation in _dataValidations)
            {
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

        /// <summary>
        /// Create a new DV with an initial area.
        /// </summary>
        internal XLDataValidation Create(XLSheetRange area)
        {
            var dv = new XLDataValidation(_worksheet);
            _dataValidations.Add(dv);
            AddArea(dv, area);
            return dv;
        }

        /// <summary>
        /// Create a new DV that is created from another DV from different sheet.
        /// </summary>
        internal XLDataValidation CopyFrom(XLDataValidation original)
        {
            var dv = new XLDataValidation(_worksheet);
            _dataValidations.Add(dv);
            dv.CopyFrom(original);
            return dv;
        }

        internal void Delete(XLSheetRange areaToDelete)
        {
            for (var i = _dataValidations.Count - 1; i >= 0; --i)
            {
                var dataValidation = _dataValidations[i];
                foreach (var dataValidationArea in dataValidation.Areas)
                {
                    if (dataValidationArea.Intersects(areaToDelete))
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
            _dataValidations.Clear();

            while (rules.Any())
            {
                var consRule = rules.First();
                _dataValidations.Add(consRule);
                var similarRules = rules.Where(r => areEqual(consRule, r)).ToList();
                similarRules.ForEach(r => rules.Remove(r));

                IXLRanges consolidatedRanges = new XLRanges(_worksheet);
                foreach (var similarRuleArea in similarRules.SelectMany(dv => dv.Areas))
                    consolidatedRanges.Add(_worksheet.Range(XLRangeAddress.FromSheetRange(_worksheet, similarRuleArea)));

                consolidatedRanges = consolidatedRanges.Consolidate();

                consRule.ClearRanges();
                consRule.AddRanges(consolidatedRanges);
            }
        }

        internal void AddArea(XLDataValidation modifiedDataValidation, XLSheetRange addedArea)
        {
            // Add an area to modifiedDV. This must be done carefully, because there can be only
            // one DV per cell. Due to this problem, the correspondence area-DV should be managed
            // by the DV collection and this method should be private. Change to private would
            // require change of DV to a nested class + separation of API object, so the method is
            // internal + exception.
            if (!_dataValidations.Contains(modifiedDataValidation))
                throw new ArgumentException("Data validation is not a data validation of this sheet.", nameof(modifiedDataValidation));

            // There can be only one DV per cell. Remove DVs from cells that should now belong
            // to the area and remove DVs without any cells.
            for (var i = _dataValidations.Count - 1; i >= 0; --i)
            {
                var dataValidation = _dataValidations[i];

                // Area could cover whole modifiedDataValidation and could remove the modifiedDV
                // before the addedArea could be added to the modifiedDV. To avoid this, it is not
                // cleared.
                if (dataValidation == modifiedDataValidation)
                    continue;

                dataValidation.Areas = dataValidation.Areas.DeleteWithoutShift(addedArea);
                if (dataValidation.Areas.Count == 0)
                {
                    _dataValidations.RemoveAt(i);
                }
            }

            // Ensure the modifiedDV area list contains only disjunct areas to ensure
            // the "one DV per cell" invariant.
            modifiedDataValidation.Areas = modifiedDataValidation.Areas.DeleteWithoutShift(addedArea).With(addedArea);
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
