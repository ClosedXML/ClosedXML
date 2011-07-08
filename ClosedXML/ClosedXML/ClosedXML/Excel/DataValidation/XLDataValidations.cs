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
            return _dataValidations.Any(dv => dv.Ranges.Count == 1 && dv.Ranges.Contains(range));
        }

        #endregion

        public void Delete(IXLDataValidation dataValidation)
        {
            _dataValidations.RemoveAll(dv => dv.Ranges.Equals(dataValidation.Ranges));
        }
    }
}