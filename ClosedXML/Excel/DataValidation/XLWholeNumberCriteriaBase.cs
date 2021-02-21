// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public abstract class XLWholeNumberCriteriaBase : XLValidationCriteria
    {
        protected XLWholeNumberCriteriaBase(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void Between(Int32 minValue, Int32 maxValue) => base.Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void EqualOrGreaterThan(Int32 value) => base.EqualOrGreaterThan(value.ToInvariantString());

        public void EqualOrLessThan(Int32 value) => base.EqualOrLessThan(value.ToInvariantString());

        public void EqualTo(Int32 value) => base.EqualTo(value.ToInvariantString());

        public void GreaterThan(Int32 value) => base.GreaterThan(value.ToInvariantString());

        public void LessThan(Int32 value) => base.LessThan(value.ToInvariantString());

        public void NotBetween(Int32 minValue, Int32 maxValue) => base.NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void NotEqualTo(Int32 value) => base.NotEqualTo(value.ToInvariantString());
    }
}
