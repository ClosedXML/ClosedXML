using ClosedXML.Excel.CalcEngine.Exceptions;
using OneOf;
using System;
using System.Globalization;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        private readonly XLWorkbook _workbook;
        private readonly XLWorksheet _worksheet;
        private readonly IXLAddress _formulaAddress;

        public CalcContext(CultureInfo culture, XLWorksheet worksheet)
        {
            _worksheet = worksheet;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLWorkbook workbook, XLWorksheet worksheet, IXLAddress formulaAddress)
        {
            CalcEngine = calcEngine;
            _workbook = workbook;
            _worksheet = worksheet;
            _formulaAddress = formulaAddress;
            Culture = culture;
            Converter = new ValueConverter(culture);
        }

        // TODO: Remove once legacy functions are migrated
        internal XLCalcEngine CalcEngine { get; }

        /// <summary>
        /// Worksheet of the cell the formula is calculating.
        /// </summary>
        public XLWorkbook Workbook => _workbook ?? throw new MissingContextException();

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
            var cell = worksheet.GetCell(rowNumber, columnNumber);
            if (cell is null)
                return ScalarValue.FromT1(new Number1(0));

            if (cell.IsEvaluating)
                throw new InvalidOperationException("Circular reference");

            var value = cell.Value;
            return value switch
            {
                bool logical => ScalarValue.FromT0(new Logical(logical)),
                double number => ScalarValue.FromT1(new Number1(number)),
                string text => text == string.Empty
                    ? null
                    : ScalarValue.FromT2(new Text(text)),
                DateTime date => ScalarValue.FromT1(new Number1(date.ToOADate())),
                _ => throw new NotImplementedException($"Not sure how to get error from a cell (type {value?.GetType().Name}, value {value}).")
            };
        }
    }
}
