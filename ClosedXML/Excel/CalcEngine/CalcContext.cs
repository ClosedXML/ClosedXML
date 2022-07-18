using ClosedXML.Excel.CalcEngine.Exceptions;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        private readonly XLWorksheet _worksheet;
        private readonly IXLAddress _formulaAddress;

        public CalcContext(CultureInfo culture, XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        public CalcContext(CultureInfo culture, XLWorksheet worksheet, IXLAddress formulaAddress)
        {
            _worksheet = worksheet;
            _formulaAddress = formulaAddress;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public XLWorksheet Worksheet => _worksheet ?? throw new MissingContextException();

        /// <summary>
        /// Address of the calculated formula.
        /// </summary>
        public IXLAddress FormulaAddress => _formulaAddress ?? throw new MissingContextException();

        public ValueConverter Converter { get; }

        /// <summary>
        /// A culture used for comparisons and conversions (e.g. text to number).
        /// </summary>
        public CultureInfo Culture { get; }

        /// <summary>
        /// Excel 2016 and earlier doesn't support dynamic array formulas (it used an array formulas instead). As a consequence,
        /// all arguments for scalar functions where passed through implicit intersection before calling the function.
        /// </summary>
        public bool UseImplicitIntersection => true;
    }
}
