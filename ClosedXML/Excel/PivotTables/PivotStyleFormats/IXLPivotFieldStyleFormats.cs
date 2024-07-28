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

    /// <summary>
    /// Get the style of the pivot field header. The head usually contains a name of the field.
    /// In some layouts, header is not individually displayed (e.g. compact), while in others
    /// it is (e.g. tabular).
    /// </summary>
        IXLPivotStyleFormat Header { get; }

    /// <summary>
    /// Get the style of the pivot field label values on horizontal or vertical axis.
    /// </summary>
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
