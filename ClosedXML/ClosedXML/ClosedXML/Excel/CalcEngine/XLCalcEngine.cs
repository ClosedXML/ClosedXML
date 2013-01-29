using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCalcEngine : CalcEngine
    {
        private readonly IXLWorksheet _ws;
        private readonly XLWorkbook _wb;
        public XLCalcEngine()
        {}
        public XLCalcEngine(XLWorkbook wb)
        {
            _wb = wb;
            IdentifierChars = "$:!";
        }
        public XLCalcEngine(IXLWorksheet ws): this(ws.Workbook)
        {
            _ws = ws;
        }

        public override object GetExternalObject(string identifier)
        {
            if (identifier.Contains("!") && _wb != null)
            {
                var wsName = identifier.Substring(0, identifier.IndexOf("!"));
                return new CellRangeReference(_wb.Worksheet(wsName).Range(identifier.Substring(identifier.IndexOf("!") + 1)), this);
            }

            if (_ws != null)
                return new CellRangeReference(_ws.Range(identifier), this);

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
            return _range.Cells().Select(GetValue).GetEnumerator();
        }

        private Boolean _evaluating;

        // ** implementation
        object GetValue(IXLCell cell)
        {
            if (_evaluating)
            {
                throw new Exception("Circular Reference");
            }
            try
            {
                _evaluating = true;
                var f = cell.FormulaA1;
                if (XLHelper.IsNullOrWhiteSpace(f))
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
