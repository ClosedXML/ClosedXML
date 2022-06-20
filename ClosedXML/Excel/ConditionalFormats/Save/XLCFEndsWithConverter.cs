using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFEndsWithConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var val = cf.Values[1].Value;
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            {
                conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];
            }

            conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.EndsWith;
            conditionalFormattingRule.Text = val;

            var formula = new Formula { Text = "RIGHT(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "," + val.Length.ToString() + ")=\"" + val + "\"" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }


    }
}
