using DocumentFormat.OpenXml.Office2010.Excel;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCFConvertersExtension
    {
        private readonly static Dictionary<XLConditionalFormatType, IXLCFConverterExtension> Converters;

        static XLCFConvertersExtension()
        {
            Converters = new Dictionary<XLConditionalFormatType, IXLCFConverterExtension>()
            {
                { XLConditionalFormatType.DataBar, new XLCFDataBarConverterExtension() }
            };
        }

        public XLCFConvertersExtension()
        {
        }

        public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, XLWorkbook.SaveContext context)
        {
            return Converters[conditionalFormat.ConditionalFormatType].Convert(conditionalFormat, context);
        }
    }
}