using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLDataValidations: IXLDataValidations
    {
        private List<IXLDataValidation> dataValidations = new List<IXLDataValidation>();
        public void Add(IXLDataValidation dataValidation)
        {
            dataValidations.Add(dataValidation);
        }

        public void Delete(IXLDataValidation dataValidation)
        {
            dataValidations.RemoveAll(dv=>dv.Ranges.Equals(dataValidation.Ranges));
        }

        public IEnumerator<IXLDataValidation> GetEnumerator()
        {
            return dataValidations.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean ContainsSingle(IXLRange range)
        {
            foreach (var dv in dataValidations)
            {
                if (dv.Ranges.Count == 1 && dv.Ranges.Contains(range))
                    return true;
            }
            return false;
        }

        
    }
}
