using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal class XLCFCellIsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = GetQuoted(cf.Values[1]);

            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
                conditionalFormattingRule.FormatId = (UInt32)context.DifferentialFormats[cfStyle];

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

        private String GetQuoted(XLFormula formula)
        {
            String value = formula.Value;

            if (formula.IsFormula ||
                value.StartsWith("\"") && value.EndsWith("\"") ||
                Double.TryParse(value, XLHelper.NumberStyle, XLHelper.ParseCulture, out double num))
            {
                return value;
            }

            return String.Format("\"{0}\"", value.Replace("\"", "\"\""));
        }
    }
}
