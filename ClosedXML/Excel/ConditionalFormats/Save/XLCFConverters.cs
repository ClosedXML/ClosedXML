using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

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
                {XLConditionalFormatType.IconSet, new XLCFIconSetConverter()}
            };
        }
        public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, Int32 priority, XLWorkbook.SaveContext context)
        {
            return Converters[conditionalFormat.ConditionalFormatType].Convert(conditionalFormat, priority, context);
        }
    }
}
