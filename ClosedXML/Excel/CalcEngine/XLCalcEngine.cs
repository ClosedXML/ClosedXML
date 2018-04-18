using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCalcEngine : CalcEngine
    {
        private readonly IXLWorksheet _ws;
        private readonly XLWorkbook _wb;

        public XLCalcEngine()
        { }

        public XLCalcEngine(XLWorkbook wb)
        {
            _wb = wb;
            IdentifierChars = new char[] { '$', ':', '!' };
        }

        public XLCalcEngine(IXLWorksheet ws) : this(ws.Workbook)
        {
            _ws = ws;
        }

        private IList<IXLRange> _cellRanges;
        /// <summary>
        /// Get a collection of cell ranges included into the expression. Order is not preserved.
        /// </summary>
        /// <param name="expression">Formula to parse.</param>
        /// <returns>Collection of ranges included into the expression.</returns>
        public IEnumerable<IXLRange> GetAffectingRanges(string expression)
        {
            _cellRanges = new List<IXLRange>();
            Parse(expression);
            var ranges = _cellRanges;
            _cellRanges = null;
            var visitedRanges = new HashSet<IXLRangeAddress>(new XLRangeAddressComparer(true));
            foreach (var range in ranges)
            {
                if (!visitedRanges.Contains(range.RangeAddress))
                {
                    visitedRanges.Add(range.RangeAddress);
                    yield return range;
                }
            }
        }

        public IEnumerable<IXLCell> GetAffectingCells(string expression)
        {
            if (String.IsNullOrWhiteSpace(expression) && String.IsNullOrEmpty(expression))
                yield break;
            else
            {
                var ranges = GetAffectingRanges(expression);
                var visitedCells = new HashSet<IXLAddress>(new XLAddressComparer(true));
                var cells = ranges.SelectMany(range => range.Cells()).Distinct();
                foreach (var cell in cells)
                {
                    if (!visitedCells.Contains(cell.Address))
                    {
                        visitedCells.Add(cell.Address);
                        yield return cell;
                    }
                }
            }
        }

        public override object GetExternalObject(string identifier)
        {
            if (identifier.Contains("!") && _wb != null)
            {
                var referencedSheetNames = identifier.Split(':')
                    .Select(part =>
                    {
                        if (part.Contains("!"))
                            return part.Substring(0, part.IndexOf('!')).ToLower();
                        else
                            return null;
                    })
                    .Where(sheet => sheet != null)
                    .Distinct();

                if (!referencedSheetNames.Any())
                    return GetCellRangeReference(_ws.Range(identifier));
                else if (referencedSheetNames.Count() > 1)
                    throw new ArgumentOutOfRangeException(referencedSheetNames.Last(), "Cross worksheet references may references no more than 1 other worksheet");
                else
                {
                    IXLWorksheet worksheet;
                    if (!_wb.TryGetWorksheet(referencedSheetNames.Single(), out worksheet))
                        throw new ArgumentOutOfRangeException(referencedSheetNames.Single(), "The required worksheet cannot be found");

                    identifier = identifier.ToLower().Replace(string.Format("{0}!", worksheet.Name.ToLower()), "");

                    return GetCellRangeReference(worksheet.Range(identifier));
                }
            }
            else if (_ws != null)
                return GetCellRangeReference(_ws.Range(identifier));
            else
                return identifier;
        }

        private CellRangeReference GetCellRangeReference(IXLRange range)
        {
            var res = new CellRangeReference(range, this);
            if (_cellRanges != null)
                _cellRanges.Add(res.Range);
            return res;
        }

        //TODO Make a separate internal class?
        private class XLRangeAddressComparer : IEqualityComparer<IXLRangeAddress>
        {
            private bool _ignoreFixed;
            private XLAddressComparer _addressComparer;
            public XLRangeAddressComparer(bool ignoreFixed)
            {
                _ignoreFixed = ignoreFixed;
                _addressComparer = new XLAddressComparer(_ignoreFixed);
            }

            public bool Equals(IXLRangeAddress x, IXLRangeAddress y)
            {
                return (x == null && y == null) ||
                    (x != null && y != null &&
                    _addressComparer.Equals(x.FirstAddress, y.FirstAddress) &&
                    _addressComparer.Equals(x.LastAddress, y.LastAddress));
            }

            public int GetHashCode(IXLRangeAddress obj)
            {
                return new
                {
                    FirstHash = _addressComparer.GetHashCode(obj.FirstAddress),
                    LastHash = _addressComparer.GetHashCode(obj.LastAddress),
                }.GetHashCode();
            }
        }

        //TODO Make a separate internal class?
        private class XLAddressComparer : IEqualityComparer<IXLAddress>
        {
            private bool _ignoreFixed;
            public XLAddressComparer(bool ignoreFixed)
            {
                _ignoreFixed = ignoreFixed;
            }

            public bool Equals(IXLAddress x, IXLAddress y)
            {
                return (x == null && y == null) ||
                    (x != null && y != null &&
                    string.Equals(x.Worksheet.Name, y.Worksheet.Name, StringComparison.InvariantCultureIgnoreCase) &&
                    x.ColumnNumber == y.ColumnNumber &&
                    x.RowNumber == y.RowNumber &&
                    (_ignoreFixed || x.FixedColumn == y.FixedColumn &&
                                     x.FixedRow == y.FixedRow));
            }

            public int GetHashCode(IXLAddress obj)
            {
                return new {
                    WorksheetName = obj.Worksheet.Name.ToUpperInvariant(),
                    obj.ColumnNumber,
                    obj.RowNumber,
                    FixedColumn = (_ignoreFixed ? false : obj.FixedColumn),
                    FixedRow = (_ignoreFixed ? false : obj.FixedRow)
                }.GetHashCode();
            }
        }
    }

    internal class CellRangeReference : IValueObject, IEnumerable
    {
        private IXLRange _range;
        private XLCalcEngine _ce;

        public CellRangeReference(IXLRange range, XLCalcEngine ce)
        {
            _range = range;
            _ce = ce;
        }

        public IXLRange Range { get { return _range; } }

        // ** IValueObject
        public object GetValue()
        {
            return GetValue(_range.FirstCell());
        }

        // ** IEnumerable
        public IEnumerator GetEnumerator()
        {
            var maxRow = Math.Min(_range.RangeAddress.LastAddress.RowNumber, _range.Worksheet.LastCellUsed().Address.RowNumber);
            var maxCol = Math.Min(_range.RangeAddress.LastAddress.ColumnNumber, _range.Worksheet.LastCellUsed().Address.ColumnNumber);
            using (var trimmedRange = (XLRangeBase)_range.Worksheet.Range(_range.FirstCell().Address, new XLAddress(maxRow, maxCol, false, false)))
                return trimmedRange.CellValues().GetEnumerator();
        }

        private Boolean _evaluating;

        // ** implementation
        private object GetValue(IXLCell cell)
        {
            if (_evaluating || (cell as XLCell).IsEvaluating)
            {
                throw new InvalidOperationException($"Circular Reference occured during evaluation. Cell: {cell.Address.ToString(XLReferenceStyle.Default, true)}");
            }
            try
            {
                _evaluating = true;
                var f = cell.FormulaA1;
                if (String.IsNullOrWhiteSpace(f))
                    return cell.Value;
                else
                {
                    return (cell as XLCell).Evaluate();
                }
            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
