using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        public CalcContext(IXLWorksheet worksheet, CultureInfo culture)
        {
            Worksheet = worksheet;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public IXLWorksheet Worksheet { get; }

        public ValueConverter Converter { get; }

        /// <summary>
        /// A culture used for comparisons and conversions (e.g. text to number).
        /// </summary>
        public CultureInfo Culture { get; }
    }
}
