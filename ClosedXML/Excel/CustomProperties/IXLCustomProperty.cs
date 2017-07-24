using System;

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
