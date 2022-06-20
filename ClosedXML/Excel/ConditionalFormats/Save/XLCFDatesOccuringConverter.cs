using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCFDatesOccurringConverter : IXLCFConverter
    {
        private static readonly IDictionary<XLTimePeriod, string> formulaTemplates = new Dictionary<XLTimePeriod, string>()
        {
            [XLTimePeriod.Today] = "FLOOR({0},1)=TODAY()",
            [XLTimePeriod.Yesterday] = "FLOOR({0},1)=TODAY()-1",
            [XLTimePeriod.Tomorrow] = "FLOOR({0},1)=TODAY()+1",
            [XLTimePeriod.InTheLast7Days] = "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
            [XLTimePeriod.ThisMonth] = "AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))",
            [XLTimePeriod.LastMonth] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))",
            [XLTimePeriod.NextMonth] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)),YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
            [XLTimePeriod.ThisWeek] = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            [XLTimePeriod.LastWeek] = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
            [XLTimePeriod.NextWeek] = "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))"
        };

        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
            var cfStyle = (cf.Style as XLStyle).Value;
            if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            {
                conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];
            }

            conditionalFormattingRule.TimePeriod = cf.TimePeriod.ToOpenXml();

            var address = cf.Range.RangeAddress.FirstAddress.ToStringRelative(false);
            var formula = new Formula { Text = string.Format(formulaTemplates[cf.TimePeriod], address) };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
