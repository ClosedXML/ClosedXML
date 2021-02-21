// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public class XLDecimalCriteria : XLValidationCriteria
    {
        public XLDecimalCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void Between(Double minValue, Double maxValue) => base.Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void EqualOrGreaterThan(Double value) => base.EqualOrGreaterThan(value.ToInvariantString());

        public void EqualOrLessThan(Double value) => base.EqualOrLessThan(value.ToInvariantString());

        public void EqualTo(Double value) => base.EqualTo(value.ToInvariantString());

        public void GreaterThan(Double value) => base.GreaterThan(value.ToInvariantString());

        public void LessThan(Double value) => base.LessThan(value.ToInvariantString());

        public void NotBetween(Double minValue, Double maxValue) => base.NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void NotEqualTo(Double value) => base.NotEqualTo(value.ToInvariantString());
    }
}
