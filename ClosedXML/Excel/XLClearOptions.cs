using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLClearOptions
    {
        Contents                = 1 << 0,
        DataType                = 1 << 1,
        NormalFormats           = 1 << 2,
        ConditionalFormats      = 1 << 3,
        Comments                = 1 << 4,
        DataValidation          = 1 << 5,

        AllFormats = NormalFormats | ConditionalFormats,
        All = Contents | DataType | NormalFormats | ConditionalFormats | Comments | DataValidation
    }
}
