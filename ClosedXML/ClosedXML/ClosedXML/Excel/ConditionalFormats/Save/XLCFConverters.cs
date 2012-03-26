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
            Converters.Add(XLConditionalFormatType.BeginsWith, new XLCFBeginsWithConverter());
        }
        public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, Int32 priority, XLWorkbook.SaveContext context)
        {
            return Converters[conditionalFormat.ConditionalFormatType].Convert(conditionalFormat, priority, context);
        }
    }
}
