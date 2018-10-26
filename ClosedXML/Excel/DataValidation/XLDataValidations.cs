using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    using System.Collections;
    using System.Linq;

    internal class XLDataValidations : IXLDataValidations
    {
        private readonly List<IXLDataValidation> _dataValidations = new List<IXLDataValidation>();

        #region IXLDataValidations Members

        public void Add(IXLDataValidation dataValidation)
        {
            _dataValidations.Add(dataValidation);
        }

        public void Delete(Predicate<IXLDataValidation> predicate)
        {
            _dataValidations.RemoveAll(predicate);
        }

        public IEnumerator<IXLDataValidation> GetEnumerator()
        {
            return _dataValidations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
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

        #endregion IXLDataValidations Members

        public void Delete(IXLDataValidation dataValidation)
        {
            _dataValidations.RemoveAll(dv => dv.Ranges.Equals(dataValidation.Ranges));
        }

        public void Delete(IXLRange range)
        {
            _dataValidations.RemoveAll(dv => dv.Ranges.Contains(range));
        }

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
                    dv1.MaxValue == dv2.MaxValue;
            };

            var rules = _dataValidations.ToList();
            _dataValidations.Clear();

            while (rules.Any())
            {
                var similarRules = rules.Where(r => areEqual(rules.First(), r)).ToList();
                similarRules.ForEach(r => rules.Remove(r));

                var consRule = similarRules.First();
                var ranges = similarRules.SelectMany(dv => dv.Ranges).ToList();
                consRule.Ranges.RemoveAll();
                ranges.ForEach(r => consRule.Ranges.Add(r));
                consRule.Ranges = consRule.Ranges.Consolidate();
                _dataValidations.Add(consRule);
            }
        }
    }
}
