using ClosedXML.Excel.CalcEngine.Exceptions;
using OneOf;
using System;
using System.Globalization;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Range>;

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

        internal ScalarValue GetCellValue(XLWorksheet worksheet, int rowNumber, int columnNumber)
        {
            return GetCellValueOrBlank(worksheet, rowNumber, columnNumber) ?? ScalarValue.FromT1(new Number1(0));
        }

        internal ScalarValue? GetCellValueOrBlank(XLWorksheet worksheet, int rowNumber, int columnNumber)
        {
            worksheet ??= _worksheet;
            var value = worksheet.GetCellValue(rowNumber, columnNumber);
            return value switch
            {
                bool logical => ScalarValue.FromT0(new Logical(logical)),
                double number => ScalarValue.FromT1(new Number1(number)),
                string text => text == string.Empty
                    ? null
                    : ScalarValue.FromT2(new Text(text)),
                _ => throw new NotImplementedException("Not sure how to get error from a cell.")
            };
        }

    }
}
