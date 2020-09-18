using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCFConverters
    {
        private static readonly Dictionary<XLConditionalFormatType, IXLCFConverter> Converters;
        static XLCFConverters()
        {
            Converters = new Dictionary<XLConditionalFormatType, IXLCFConverter>
            {
                {XLConditionalFormatType.ColorScale, new XLCFColorScaleConverter()},
                {XLConditionalFormatType.StartsWith, new XLCFStartsWithConverter()},
                {XLConditionalFormatType.EndsWith, new XLCFEndsWithConverter()},
                {XLConditionalFormatType.IsBlank, new XLCFIsBlankConverter()},
                {XLConditionalFormatType.NotBlank, new XLCFNotBlankConverter()},
                {XLConditionalFormatType.IsError, new XLCFIsErrorConverter()},
                {XLConditionalFormatType.NotError, new XLCFNotErrorConverter()},
                {XLConditionalFormatType.ContainsText, new XLCFContainsConverter()},
                {XLConditionalFormatType.NotContainsText, new XLCFNotContainsConverter()},
                {XLConditionalFormatType.CellIs, new XLCFCellIsConverter()},
                {XLConditionalFormatType.IsUnique, new XLCFUniqueConverter()},
                {XLConditionalFormatType.IsDuplicate, new XLCFUniqueConverter()},
                {XLConditionalFormatType.Expression, new XLCFCellIsConverter()},
                {XLConditionalFormatType.Top10, new XLCFTopConverter()},
                {XLConditionalFormatType.DataBar, new XLCFDataBarConverter()},
                {XLConditionalFormatType.IconSet, new XLCFIconSetConverter()},
                {XLConditionalFormatType.TimePeriod, new XLCFDatesOccurringConverter()}
            };
        }

        public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, Int32 priority, XLWorkbook.SaveContext context)
        {
            if (!Converters.TryGetValue(conditionalFormat.ConditionalFormatType, out var converter))
                throw new NotImplementedException(string.Format("Conditional formatting rule '{0}' hasn't been implemented", conditionalFormat.ConditionalFormatType));

            return converter.Convert(conditionalFormat, priority, context);
        }
    }
}
