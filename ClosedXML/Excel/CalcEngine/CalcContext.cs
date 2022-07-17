using ClosedXML.Excel.CalcEngine.Exceptions;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        private readonly IXLWorksheet _worksheet;

        public CalcContext(IXLWorksheet worksheet, CultureInfo culture)
        {
            _worksheet = worksheet;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public IXLWorksheet Worksheet => _worksheet ?? throw new MissingContextException();

        /// <summary>
        /// Address of the calculated formula.
        /// </summary>
        public IXLAddress FormulaAddress => throw new MissingContextException();

        public ValueConverter Converter { get; }

        /// <summary>
        /// A culture used for comparisons and conversions (e.g. text to number).
        /// </summary>
        public CultureInfo Culture { get; }
    }
}
