using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCustomProperty: IXLCustomProperty
    {
        XLWorkbook workbook;
        public XLCustomProperty(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }
        private String name;
        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                if (workbook.CustomProperties.Any(t => t.Name == value))
                    throw new ArgumentException(String.Format("This workbook already contains a custom property named '{0}'", value));

                name = value;
            }
        }
        public XLCustomPropertyType Type 
        {
            get
            {
                Double dTest;
                if (Value is DateTime)
                    return XLCustomPropertyType.Date;
                else if (Value is Boolean)
                    return XLCustomPropertyType.Boolean;
                else if (Double.TryParse(Value.ToString(), out dTest))
                    return XLCustomPropertyType.Number;
                else
                    return XLCustomPropertyType.Text;
            }
        }
        public Object Value { get; set; }
        public T GetValue<T>()
        {
            return (T)Convert.ChangeType(Value, typeof(T));
        }
    }
}
