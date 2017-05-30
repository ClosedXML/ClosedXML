using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFCellIsConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            String val = GetQuoted(cf.Values[1]);


            var conditionalFormattingRule = new ConditionalFormattingRule { FormatId = (UInt32)context.DifferentialFormats[cf.Style], Operator = cf.Operator.ToOpenXml(), Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var formula = new Formula();
            if (cf.Operator == XLCFOperator.Equal || cf.Operator == XLCFOperator.NotEqual)
                formula.Text = val;
            else
                formula.Text = val;
            conditionalFormattingRule.Append(formula);

            if(cf.Operator == XLCFOperator.Between || cf.Operator == XLCFOperator.NotBetween)
            {
                var formula2 = new Formula { Text = GetQuoted(cf.Values[2]) };
                conditionalFormattingRule.Append(formula2);    
            }

            return conditionalFormattingRule;
        }

        private String GetQuoted(XLFormula formula)
        {
            String value = formula.Value;
            Double num;
            if ((!Double.TryParse(value, out num) && !formula.IsFormula) && value[0] != '\"' && !value.EndsWith("\""))
                return String.Format("\"{0}\"", value.Replace("\"", "\"\""));

            return value;
        }


    }
}
