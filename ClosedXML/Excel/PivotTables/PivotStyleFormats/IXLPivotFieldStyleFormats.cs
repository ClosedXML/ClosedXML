#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    public interface IXLPivotFieldStyleFormats
    {
        IXLPivotValueStyleFormat DataValuesFormat { get; }
        IXLPivotStyleFormat Header { get; }
        IXLPivotStyleFormat Label { get; }
        IXLPivotStyleFormat Subtotal { get; }
    }

    // TODO: Implement
    internal class XLPivotTableAxisFieldStyleFormats : IXLPivotFieldStyleFormats
    {
        public IXLPivotValueStyleFormat DataValuesFormat { get; }
        public IXLPivotStyleFormat Header { get; }
        public IXLPivotStyleFormat Label { get; }
        public IXLPivotStyleFormat Subtotal { get; }
    }
}
