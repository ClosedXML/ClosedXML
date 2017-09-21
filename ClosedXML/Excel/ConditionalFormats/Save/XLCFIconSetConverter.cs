using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFIconSetConverter:IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, Int32 priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

            var iconSet = new IconSet {ShowValue = !cf.ShowIconOnly, Reverse = cf.ReverseIconOrder, IconSetValue = cf.IconSetStyle.ToOpenXml()};
            Int32 count = cf.Values.Count;
            for(Int32 i=1;i<= count; i++ )
            {
                var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = cf.ContentTypes[i].ToOpenXml(), Val = cf.Values[i].Value, GreaterThanOrEqual = cf.IconSetOperators[i] == XLCFIconSetOperator.EqualOrGreaterThan};    
                iconSet.Append(conditionalFormatValueObject);
                
            }
            conditionalFormattingRule.Append(iconSet);
            return conditionalFormattingRule;
        }
    }
}
