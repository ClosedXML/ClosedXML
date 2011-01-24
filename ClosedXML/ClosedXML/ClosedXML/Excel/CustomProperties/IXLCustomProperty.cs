using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLCustomPropertyType { Text, Number, Date, Boolean}
    public interface IXLCustomProperty
    {
        String Name { get; set; }
        XLCustomPropertyType Type { get; }
        Object Value { get; set; }
        T GetValue<T>();
    }
}
