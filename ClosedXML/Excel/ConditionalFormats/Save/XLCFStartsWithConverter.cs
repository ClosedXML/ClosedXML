using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFStartsWithConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            if (!cf.Style.Equals(XLWorkbook.DefaultStyle))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cf.Style];

            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.BeginsWith;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "LEFT(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "," + val.Length.ToString() + ")=\"" + val + "\"" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
