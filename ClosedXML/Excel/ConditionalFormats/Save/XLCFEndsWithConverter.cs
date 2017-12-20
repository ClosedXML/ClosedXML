using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFEndsWithConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = cf.Values[1].Value;
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            if (!cf.Style.Equals(XLWorkbook.DefaultStyle))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cf.Style];

            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.EndsWith;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "RIGHT(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "," + val.Length.ToString() + ")=\"" + val + "\"" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }


    }
}
