#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel
{
    /// <summary>
    /// Interface to change the style of a <see cref="IXLPivotField"/> or it's parts.
    /// </summary>
    public interface IXLPivotFieldStyleFormats
    {
        /// <summary>
        /// Pivot table style of the field values displayed in the data area of the pivot table.
        /// </summary>
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
