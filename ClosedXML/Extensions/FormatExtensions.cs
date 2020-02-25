// Keep this file CodeMaid organised and cleaned
using ExcelNumberFormat;
using System.Globalization;

namespace ClosedXML.Extensions
{
    internal static class FormatExtensions
    {
        public static string ToExcelFormat(this object o, string format)
        {
            var nf = new NumberFormat(format);
            if (!nf.IsValid)
                return format;

            return nf.Format(o, CultureInfo.InvariantCulture);
        }
    }
}
