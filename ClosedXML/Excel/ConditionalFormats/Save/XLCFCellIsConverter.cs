using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFCellIsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var val = GetQuoted(cf.Values[1]);

            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            {
                conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];
            }

            conditionalFormattingRule.Operator = cf.Operator.ToOpenXml();

            var formula = new Formula(val);
            conditionalFormattingRule.Append(formula);

            if (cf.Operator == XLCFOperator.Between || cf.Operator == XLCFOperator.NotBetween)
            {
                var formula2 = new Formula { Text = GetQuoted(cf.Values[2]) };
                conditionalFormattingRule.Append(formula2);
            }

            return conditionalFormattingRule;
        }

        private string GetQuoted(XLFormula formula)
        {
            var value = formula.Value;

            if (formula.IsFormula ||
                value.StartsWith("\"") && value.EndsWith("\"") ||
                double.TryParse(value, XLHelper.NumberStyle, XLHelper.ParseCulture, out var num))
            {
                return value;
            }

            return string.Format("\"{0}\"", value.Replace("\"", "\"\""));
        }
    }
}
