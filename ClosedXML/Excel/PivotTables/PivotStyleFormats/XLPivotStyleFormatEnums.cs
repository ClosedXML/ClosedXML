using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLPivotStyleFormatElement
    {
        None = 0,
        Label = 1 << 1,
        Data = 1 << 2,

        All = Label | Data
    }

    internal enum XLPivotStyleFormatTarget
    {
        PivotTable,
        GrandTotal,
        Subtotal,
        Header,
        Label,
        Data
    }
}
