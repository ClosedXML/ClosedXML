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

        public XLFormula(string value)
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

        internal string _value;
        public string Value 
        { 
            get { return _value; }
            set
            {
                if (value == null)
                {
                    _value = string.Empty;
                }
                else
                {
                    _value = value.Trim();
                    IsFormula = !string.IsNullOrWhiteSpace(_value) && _value.TrimStart()[0] == '=' ;
                    if (IsFormula)
                    {
                        _value = _value.Substring(1);
                    }
                }
                

            }
        }

        public bool IsFormula { get; internal set; }
    }
}
