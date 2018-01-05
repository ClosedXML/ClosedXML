using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLCFConverters
    {
        private static readonly Dictionary<XLConditionalFormatType, IXLCFConverter> Converters;
        static XLCFConverters()
        {
            Converters = new Dictionary<XLConditionalFormatType, IXLCFConverter>();
            Converters.Add(XLConditionalFormatType.ColorScale, new XLCFColorScaleConverter());
            Converters.Add(XLConditionalFormatType.StartsWith, new XLCFStartsWithConverter());
            Converters.Add(XLConditionalFormatType.EndsWith, new XLCFEndsWithConverter());
            Converters.Add(XLConditionalFormatType.IsBlank, new XLCFIsBlankConverter());
            Converters.Add(XLConditionalFormatType.NotBlank, new XLCFNotBlankConverter());
            Converters.Add(XLConditionalFormatType.IsError, new XLCFIsErrorConverter());
            Converters.Add(XLConditionalFormatType.NotError, new XLCFNotErrorConverter());
            Converters.Add(XLConditionalFormatType.ContainsText, new XLCFContainsConverter());
            Converters.Add(XLConditionalFormatType.NotContainsText, new XLCFNotContainsConverter());
            Converters.Add(XLConditionalFormatType.CellIs, new XLCFCellIsConverter());
            Converters.Add(XLConditionalFormatType.IsUnique, new XLCFUniqueConverter());
            Converters.Add(XLConditionalFormatType.IsDuplicate, new XLCFUniqueConverter());
            Converters.Add(XLConditionalFormatType.Expression, new XLCFCellIsConverter());
            Converters.Add(XLConditionalFormatType.Top10, new XLCFTopConverter());
            Converters.Add(XLConditionalFormatType.DataBar, new XLCFDataBarConverter());
            Converters.Add(XLConditionalFormatType.IconSet, new XLCFIconSetConverter());
            Converters.Add(XLConditionalFormatType.TimePeriod, new XLCFDatesOccurringConverter());

            foreach (var cft in Enum.GetValues(typeof(XLConditionalFormatType)).Cast<XLConditionalFormatType>())
            {
                if (!Converters.ContainsKey(cft))
                    Converters.Add(cft, null);
            }
        }

        public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, Int32 priority, XLWorkbook.SaveContext context)
        {
            var converter = Converters[conditionalFormat.ConditionalFormatType];
            if (converter == null)
                throw new NotImplementedException(string.Format("Conditional formatting rule '{0}' hasn't been implemented", conditionalFormat.ConditionalFormatType));

            return converter.Convert(conditionalFormat, priority, context);
        }
    }
}
