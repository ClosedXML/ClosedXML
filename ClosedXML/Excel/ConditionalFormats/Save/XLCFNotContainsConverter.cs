using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFNotContainsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cfStyle.Key];

            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.NotContains;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "ISERROR(SEARCH(\"" + val + "\"," + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
