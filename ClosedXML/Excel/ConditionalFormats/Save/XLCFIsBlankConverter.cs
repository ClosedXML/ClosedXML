using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFIsBlankConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = cf.Style.Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyle.Value))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cfStyle.Key];

            var formula = new Formula { Text = "LEN(TRIM(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))=0" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
