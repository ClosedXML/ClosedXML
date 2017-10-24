using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFTopConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            UInt32 val = UInt32.Parse(cf.Values[1].Value);
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            if (!cf.Style.Equals(XLWorkbook.DefaultStyle))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cf.Style];

            conditionalFormattingRule.Percent = cf.Percent;
            conditionalFormattingRule.Rank = val;
            conditionalFormattingRule.Bottom = cf.Bottom;
            return conditionalFormattingRule;
        }
    }
}
