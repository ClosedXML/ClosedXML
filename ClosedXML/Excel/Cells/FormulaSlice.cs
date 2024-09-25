using System.Collections.Generic;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    internal class FormulaSlice : ISlice
    {
        private readonly XLWorksheet _sheet;
        private readonly XLCalcEngine _engine;
        private readonly Slice<XLCellFormula?> _formulas = new();

        public FormulaSlice(XLWorksheet sheet)
        {
            _sheet = sheet;
            _engine = sheet.Workbook.CalcEngine;
        }

        public bool IsEmpty => _formulas.IsEmpty;

        public int MaxColumn => _formulas.MaxColumn;

        public int MaxRow => _formulas.MaxRow;

        public Dictionary<int, int>.KeyCollection UsedColumns => _formulas.UsedColumns;

        public IEnumerable<int> UsedRows => _formulas.UsedRows;

        public void Clear(XLSheetRange range)
        {
            _formulas.Clear(range);
        }

        public void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete)
        {
            _formulas.DeleteAreaAndShiftLeft(rangeToDelete);
        }

        public void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete)
        {
            _formulas.DeleteAreaAndShiftUp(rangeToDelete);
        }

        public IEnumerator<XLSheetPoint> GetEnumerator(XLSheetRange range, bool reverse = false)
        {
            return _formulas.GetEnumerator(range, reverse);
        }

        public void InsertAreaAndShiftDown(XLSheetRange range)
        {
            _formulas.InsertAreaAndShiftDown(range);
        }

        public void InsertAreaAndShiftRight(XLSheetRange range)
        {
            _formulas.InsertAreaAndShiftRight(range);
        }

        public bool IsUsed(XLSheetPoint address)
        {
            return _formulas.IsUsed(address);
        }

        public void Swap(XLSheetPoint sp1, XLSheetPoint sp2)
        {
            var value1 = _formulas[sp1];
            var value2 = _formulas[sp2];

            value1 = value1?.GetMovedTo(sp1, sp2);
            value2 = value2?.GetMovedTo(sp2, sp1);

            Set(sp1, value2);
            Set(sp2, value1);
        }

        internal XLCellFormula? Get(XLSheetPoint point)
        {
            return _formulas[point];
        }

        internal void Set(XLSheetPoint point, XLCellFormula? formula)
        {
            // Can't ref, because it is an alias for a memory and thus wouldn't hold old formula.
            var original = _formulas[point];
            if (ReferenceEquals(original, formula))
                return;

            _formulas.Set(point, formula);

            // Remove first, so calc chain doesn't choke on two formulas
            // in one cell when changing a formula of a cell.
            var bookPoint = new XLBookPoint(_sheet.SheetId, point);
            if (original is not null)
                _engine.RemoveFormula(bookPoint, original);

            if (formula is not null)
                _engine.AddNormalFormula(bookPoint, _sheet.Name, formula, _sheet.Workbook);
        }

        /// <summary>
        /// Set all cells in a <paramref name="range"/> to the array formula.
        /// </summary>
        /// <remarks>
        /// This method doesn't check that formula doesn't damage other array formulas.
        /// </remarks>
        internal void SetArray(XLSheetRange range, XLCellFormula? arrayFormula)
        {
            for (var row = range.TopRow; row <= range.BottomRow; ++row)
            {
                for (var col = range.LeftColumn; col <= range.RightColumn; ++col)
                {
                    var point = new XLSheetPoint(row, col);
                    var original = _formulas[point];

                    _formulas.Set(point, arrayFormula);

                    // The formula removal removes formula from dependency tree
                    // (number of cells formula affects doesn't matter) and also
                    // removes point from the calc chain. Therefore, it works for
                    // array and normal formulas.
                    var bookPoint = new XLBookPoint(_sheet.SheetId, point);
                    if (original is not null)
                        _engine.RemoveFormula(bookPoint, original);
                }
            }

            if (arrayFormula is not null)
                _engine.AddArrayFormula(range, arrayFormula, _sheet);
        }

        internal Slice<XLCellFormula>.Enumerator GetForwardEnumerator(XLSheetRange range)
        {
            return new Slice<XLCellFormula>.Enumerator(_formulas!, range);
        }

        /// <summary>
        /// Mark all formulas in a range as dirty.
        /// </summary>
        internal void MarkDirty(XLSheetRange range)
        {
            using var enumerator = GetForwardEnumerator(range);
            while (enumerator.MoveNext())
            {
                enumerator.Current.IsDirty = true;
            }
        }
    }
}
