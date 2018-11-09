using System;

namespace ClosedXML.Excel
{
    [Flags]
    public enum XLCellsUsedOptions
    {
        Contents                = 1 << 0,
        DataType                = 1 << 1,
        NormalFormats           = 1 << 2,
        ConditionalFormats      = 1 << 3,
        Comments                = 1 << 4,
        DataValidation          = 1 << 5,
        MergedRanges            = 1 << 6,

        AllFormats = NormalFormats | ConditionalFormats,
        AllContents = Contents | DataType | Comments | MergedRanges,
        All = Contents | DataType | NormalFormats | ConditionalFormats | Comments | DataValidation | MergedRanges
    }

    internal static class XLCellsUsedOptionsExtensions
    {
        public static XLClearOptions ToClearOptions(this XLCellsUsedOptions options)
        {
            return (XLClearOptions)options;
        }
    }
}
