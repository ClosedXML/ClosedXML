using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDataValidations: IXLDataValidations
    {
        private Dictionary<String, IXLDataValidation> dataValidations = new Dictionary<String, IXLDataValidation>();
        public void Add(IXLDataValidation dataValidation)
        {
            dataValidations.Add(dataValidation.Range.RangeAddress.ToString(), dataValidation);
        }

        public void Delete(IXLDataValidation dataValidation)
        {
            dataValidations.Remove(dataValidation.Range.RangeAddress.ToString());
        }

        public IEnumerator<IXLDataValidation> GetEnumerator()
        {
            return dataValidations.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
