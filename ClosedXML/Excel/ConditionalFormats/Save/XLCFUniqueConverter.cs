using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFUniqueConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = cf.Style.Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyle.Value))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cfStyle.Key];

            return conditionalFormattingRule;
        }
    }
}
