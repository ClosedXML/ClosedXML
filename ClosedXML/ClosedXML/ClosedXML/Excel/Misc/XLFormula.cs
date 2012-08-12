using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLFormula
    {
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
                    Double num;
                    IsFormula = !XLHelper.IsNullOrWhiteSpace(_value) && _value.TrimStart()[0] == '=' ;
                    if (IsFormula)
                        _value = _value.Substring(1);
                    //else if (!XLHelper.IsNullOrWhiteSpace(_value) && (!Double.TryParse(_value, out num) && _value[0] != '\"' && !_value.EndsWith("\"")))
                    //    _value = String.Format("\"{0}\"", _value.Replace("\"", "\"\""));
                }
                

            }
        }

        public Boolean IsFormula { get; private set; }
    }
}
