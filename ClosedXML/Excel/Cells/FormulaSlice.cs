using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class FormulaSlice : ISlice
    {
        private readonly Slice<XLCellFormula?> _formulas = new();

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
            _formulas.Swap(sp1, sp2);
        }

        internal XLCellFormula? Get(XLSheetPoint point)
        {
            return _formulas[point];
        }

        internal void Set(XLSheetPoint point, XLCellFormula? formula)
        {
            ref readonly var original = ref _formulas[point];
            if (ReferenceEquals(original, formula))
                return;

            _formulas.Set(point, formula);
        }
    }
}
