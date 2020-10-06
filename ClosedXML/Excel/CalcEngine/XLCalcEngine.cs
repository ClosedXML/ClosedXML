using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCalcEngine : CalcEngine
    {
        private static readonly XLCellByAddressComparer xlCellByAddressComparer = new XLCellByAddressComparer();
        private static readonly XLRangeByAddressComparer xlRangeByAddressComparer = new XLRangeByAddressComparer();
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
        public IEnumerable<IXLRange> GetPrecedentRanges(string expression)
        {
            _cellRanges = new List<IXLRange>();
            Parse(expression);
            return _cellRanges.Distinct(xlRangeByAddressComparer);
        }

        public IEnumerable<IXLCell> GetPrecedentCells(string expression)
        {
            if (!String.IsNullOrWhiteSpace(expression))
            {
                var ranges = GetPrecedentRanges(expression);
                return ranges.SelectMany(range => range.Cells()).Distinct(xlCellByAddressComparer);
            }

            return Enumerable.Empty<IXLCell>();
        }

        public override object GetExternalObject(string identifier)
        {
            if (identifier.Contains("!") && _wb != null)
            {
                var referencedSheetNames = identifier.Split(':')
                    .Select(part =>
                    {
                        if (part.Contains("!"))
                            return part.Substring(0, part.LastIndexOf('!')).ToLower();
                        else
                            return null;
                    })
                    .Where(sheet => sheet != null)
                    .Distinct()
                    .ToList();

                if (referencedSheetNames.Count == 0)
                    return GetCellRangeReference(_ws.Range(identifier));
                else if (referencedSheetNames.Count > 1)
                    throw new ArgumentOutOfRangeException(referencedSheetNames.Last(),
                        "Cross worksheet references may references no more than 1 other worksheet");
                else
                {
                    if (!_wb.TryGetWorksheet(referencedSheetNames.Single(), out IXLWorksheet worksheet))
                        throw new ArgumentOutOfRangeException(referencedSheetNames.Single(),
                            "The required worksheet cannot be found");

                    identifier = identifier.ToLower().Replace(string.Format("{0}!", worksheet.Name.ToLower()), "");

                    return GetCellRangeReference(worksheet.Range(identifier));
                }
            }
            else if (_ws != null)
            {
                if (TryGetNamedRange(identifier, _ws, out IXLNamedRange namedRange))
                {
                    var references = (namedRange as XLNamedRange).RangeList.Select(r =>
                        XLHelper.IsValidRangeAddress(r)
                            ? GetCellRangeReference(_ws.Workbook.Range(r))
                            : new XLCalcEngine(_ws).Evaluate(r.ToString())
                        );
                    if (references.Count() == 1)
                        return references.Single();
                    return references;
                }

                return GetCellRangeReference(_ws.Range(identifier));
            }
            else if (XLHelper.IsValidRangeAddress(identifier))
                return identifier;
            else
                return null;
        }

        private bool TryGetNamedRange(string identifier, IXLWorksheet worksheet, out IXLNamedRange namedRange)
        {
            return worksheet.NamedRanges.TryGetValue(identifier, out namedRange) ||
                   worksheet.Workbook.NamedRanges.TryGetValue(identifier, out namedRange);
        }

        private CellRangeReference GetCellRangeReference(IXLRange range)
        {
            if (range == null)
                return null;

            var res = new CellRangeReference(range, this);
            _cellRanges?.Add(res.Range);
            return res;
        }
    }
}
