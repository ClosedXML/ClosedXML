#nullable disable

using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLCellsUsedOptions
    {

        None                    = 0,
        NoConstraints           = None,

        Contents                = 1 << 0,
        DataType                = 1 << 1,
        NormalFormats           = 1 << 2,
        ConditionalFormats      = 1 << 3,
        Comments                = 1 << 4,
        DataValidation          = 1 << 5,
        MergedRanges            = 1 << 6,
        Sparklines              = 1 << 7,

        AllFormats = NormalFormats | ConditionalFormats,
        AllContents = Contents | DataType | Comments,
        All = Contents | DataType | NormalFormats | ConditionalFormats | Comments | DataValidation | MergedRanges | Sparklines
    }
}
