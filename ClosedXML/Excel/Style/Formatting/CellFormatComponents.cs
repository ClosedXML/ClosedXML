using System;

namespace ClosedXML.Excel.Formatting;

[Flags]
internal enum CellFormatComponents
{
    None = 0,
    NumberFormat = 1,
    Font = 2,
    Fill = 4,
    Border = 8,
    Alignment = 16,
    Protection = 32,
    All = NumberFormat | Font | Fill | Border | Alignment | Protection
}
