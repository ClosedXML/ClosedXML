using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFIsErrorConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = ((XLStyle)cf.Style).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
                conditionalFormattingRule.FormatId = context.GetDxfId(cfStyle);

            var formula = new Formula { Text = "ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
