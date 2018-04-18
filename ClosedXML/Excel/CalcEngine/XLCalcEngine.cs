using System;
using System.Collections;
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
                    return new CellRangeReference(_ws.Range(identifier), this);
                else if (referencedSheetNames.Count() > 1)
                    throw new ArgumentOutOfRangeException(referencedSheetNames.Last(), "Cross worksheet references may references no more than 1 other worksheet");
                else
                {
                    if (!_wb.TryGetWorksheet(referencedSheetNames.Single(), out IXLWorksheet worksheet))
                        throw new ArgumentOutOfRangeException(referencedSheetNames.Single(), "The required worksheet cannot be found");

                    identifier = identifier.ToLower().Replace(string.Format("{0}!", worksheet.Name.ToLower()), "");

                    return new CellRangeReference(worksheet.Range(identifier), this);
                }
            }
            else if (_ws != null)
                return new CellRangeReference(_ws.Range(identifier), this);
            else
                return identifier;
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
                    return new XLCalcEngine(cell.Worksheet).Evaluate(f);
            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
