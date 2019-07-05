using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLPivotStyleFormatElement
    {
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

    public enum XLPivotAreaValues
    {
        None = 0,
        Normal = 1,
        Data = 2,
        All = 3,
        Origin = 4,
        Button = 5,
        TopRight = 6,
        TopEnd = 7
    }

    public enum XLPivotTableAxisValues
    {
        AxisRow = 0,
        AxisColumn = 1,
        AxisPage = 2,
        AxisValues = 3,
    }
}
