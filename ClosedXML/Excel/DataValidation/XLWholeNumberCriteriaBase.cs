// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public abstract class XLWholeNumberCriteriaBase : XLValidationCriteria
    {
        protected XLWholeNumberCriteriaBase(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void Between(int minValue, int maxValue) => Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void EqualOrGreaterThan(int value) => EqualOrGreaterThan(value.ToInvariantString());

        public void EqualOrLessThan(int value) => EqualOrLessThan(value.ToInvariantString());

        public void EqualTo(int value) => EqualTo(value.ToInvariantString());

        public void GreaterThan(int value) => GreaterThan(value.ToInvariantString());

        public void LessThan(int value) => LessThan(value.ToInvariantString());

        public void NotBetween(int minValue, int maxValue) => NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void NotEqualTo(int value) => NotEqualTo(value.ToInvariantString());
    }
}
