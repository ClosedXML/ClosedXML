using System;

namespace ClosedXML.Excel
{
    [Flags]
    internal enum XLCellCopyOptions
    {
        None               = 0,
        Values             = 1 << 1,
        Styles             = 1 << 2,
        ConditionalFormats = 1 << 3,
        DataValidations    = 1 << 4,
        Sparklines         = 1 << 5,
        All = Values | Styles | ConditionalFormats | DataValidations | Sparklines
    }
}
