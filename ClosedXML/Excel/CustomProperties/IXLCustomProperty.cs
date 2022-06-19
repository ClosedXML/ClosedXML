namespace ClosedXML.Excel
{
    public enum XLCustomPropertyType { Text, Number, Date, Boolean}
    public interface IXLCustomProperty
    {
        string Name { get; set; }
        XLCustomPropertyType Type { get; }
        object Value { get; set; }
        T GetValue<T>();
    }
}
