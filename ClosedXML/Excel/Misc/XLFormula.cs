using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
