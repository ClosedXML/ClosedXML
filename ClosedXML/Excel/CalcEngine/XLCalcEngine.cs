using System;
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
    }
}
