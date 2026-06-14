#nullable disable

using ClosedXML.Parser;
using System;

namespace ClosedXML.Excel
{
    public class XLFormula
    {
        public XLFormula()
        {}

        public XLFormula(XLFormula defaultFormula)
        {
            _value = defaultFormula._value;
            IsFormula = defaultFormula.IsFormula;
        }

        public XLFormula(String value)
        {
            Value = value;
        }

        public XLFormula(double value)
        {
            Value = value.ToInvariantString();
        }

        public XLFormula(int value)
        {
            Value = value.ToInvariantString();
        }

        internal String _value;
        public String Value
        {
            get { return _value; }
            set
            {
                if (value == null)
                {
                    _value = String.Empty;
                }
                else
                {
                    _value = value.Trim();
                    IsFormula = !String.IsNullOrWhiteSpace(_value) && _value.TrimStart()[0] == '=' ;
                    if (IsFormula)
                        _value = _value.Substring(1);
                }


            }
        }

        public Boolean IsFormula { get; internal set; }

        internal XLFormula GetAdjustedCopy(Point sourceAnchor, Point targetAnchor)
        {
            if (!IsFormula)
                return new XLFormula(this);

            var formulaR1C1 = FormulaConverter.ToR1C1(Value, sourceAnchor.Row, sourceAnchor.Column);
            var formulaA1 = FormulaConverter.ToA1(formulaR1C1, targetAnchor.Row, targetAnchor.Column);
            return new XLFormula { _value = formulaA1, IsFormula = true };
        }
    }
}
