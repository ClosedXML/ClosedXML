﻿using ClosedXML.Excel.CalcEngine.Exceptions;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalcContext
    {
        private readonly CalcEngine _calcEngine;
        private readonly XLWorkbook _workbook;
        private readonly XLWorksheet _worksheet;
        private readonly IXLAddress _formulaAddress;

        public CalcContext(CalcEngine calcEngine, CultureInfo culture, XLCell cell)
            : this(calcEngine, culture, cell.Worksheet.Workbook, cell.Worksheet, cell.Address)
        {
        }

        public CalcContext(CalcEngine calcEngine, CultureInfo culture, XLWorkbook workbook, XLWorksheet worksheet, IXLAddress formulaAddress)
        {
            _calcEngine = calcEngine;
            _workbook = workbook;
            _worksheet = worksheet;
            _formulaAddress = formulaAddress;
            Culture = culture;
        }

        // LEGACY: Remove once legacy functions are migrated
        internal CalcEngine CalcEngine => _calcEngine ?? throw new MissingContextException();

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

        /// <summary>
        /// A culture used for comparisons and conversions (e.g. text to number).
        /// </summary>
        public CultureInfo Culture { get; }

        /// <summary>
        /// Excel 2016 and earlier doesn't support dynamic array formulas (it used an array formulas instead). As a consequence,
        /// all arguments for scalar functions where passed through implicit intersection before calling the function.
        /// </summary>
        public bool UseImplicitIntersection => true;

        /// <summary>
        /// Should functions be calculated per item of multi-values argument in the scalar parameters.
        /// </summary>
        public bool IsArrayCalculation { get; set; }

        internal ScalarValue GetCellValue(XLWorksheet worksheet, int rowNumber, int columnNumber)
        {
            worksheet ??= Worksheet;
            var cell = worksheet.GetCell(rowNumber, columnNumber);
            if (cell is null)
                return ScalarValue.Blank;

            if (cell.IsEvaluating)
                throw new InvalidOperationException($"Cell {cell.Address} is a part of circular reference.");

            return cell.Value;
        }

        /// <summary>
        /// Get cells with a value for a reference.
        /// </summary>
        /// <param name="reference">Reference for which to return cells.</param>
        /// <returns>A lazy (non-materialized) enumerable of cells with a value for the reference.</returns>
        internal IEnumerable<XLCell> GetNonBlankCells(Reference reference)
        {
            // XLCells is not suitable here, e.g. it doesn't count a cell twice if it is in multiple areas
            var nonBlankCells = Enumerable.Empty<XLCell>();
            foreach (var area in reference.Areas)
            {
                var areaCells = Worksheet.Internals.CellsCollection
                    .GetCells(
                        area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber,
                        area.LastAddress.RowNumber, area.LastAddress.ColumnNumber,
                        cell => !cell.IsEmpty());
                nonBlankCells = nonBlankCells.Concat(areaCells);
            }

            return nonBlankCells;
        }
    }
}
