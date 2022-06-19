// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    public class XLDecimalCriteria : XLValidationCriteria
    {
        public XLDecimalCriteria(IXLDataValidation dataValidation)
            : base(dataValidation)
        {
        }

        public void Between(double minValue, double maxValue) => Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void EqualOrGreaterThan(double value) => EqualOrGreaterThan(value.ToInvariantString());

        public void EqualOrLessThan(double value) => EqualOrLessThan(value.ToInvariantString());

        public void EqualTo(double value) => EqualTo(value.ToInvariantString());

        public void GreaterThan(double value) => GreaterThan(value.ToInvariantString());

        public void LessThan(double value) => LessThan(value.ToInvariantString());

        public void NotBetween(double minValue, double maxValue) => NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

        public void NotEqualTo(double value) => NotEqualTo(value.ToInvariantString());
    }
}
