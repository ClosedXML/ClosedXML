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

        #endregion

        public void Delete(IXLDataValidation dataValidation)
        {
            _dataValidations.RemoveAll(dv => dv.Ranges.Equals(dataValidation.Ranges));
        }

        public void Delete(IXLRange range)
        {
            _dataValidations.RemoveAll(dv => dv.Ranges.Contains(range));
        }
    }
}