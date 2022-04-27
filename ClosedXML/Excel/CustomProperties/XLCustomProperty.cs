using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCustomProperty : IXLCustomProperty
    {
        private readonly XLWorkbook _workbook;

        private string name;

        public XLCustomProperty(XLWorkbook workbook)
        {
            _workbook = workbook;
        }

        #region IXLCustomProperty Members

        public string Name
        {
            get { return name; }
            set
            {
                if (name == value)
                {
                    return;
                }

                if (_workbook.CustomProperties.Any(t => t.Name == value))
                {
                    throw new ArgumentException(
                        string.Format("This workbook already contains a custom property named '{0}'", value));
                }

                name = value;
            }
        }

        public XLCustomPropertyType Type
        {
            get
            {
                if (Value is DateTime)
                {
                    return XLCustomPropertyType.Date;
                }

                if (Value is bool)
                {
                    return XLCustomPropertyType.Boolean;
                }

                if (double.TryParse(Value.ToString(), out var dTest))
                {
                    return XLCustomPropertyType.Number;
                }

                return XLCustomPropertyType.Text;
            }
        }

        public object Value { get; set; }

        public T GetValue<T>()
        {
            return (T)Convert.ChangeType(Value, typeof(T));
        }

        #endregion
    }
}
