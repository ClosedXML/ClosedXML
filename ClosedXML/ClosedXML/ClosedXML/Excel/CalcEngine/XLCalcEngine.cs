using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    public class XLCalcEngine : CalcEngine
    {
        private IXLWorksheet _ws;
        public XLCalcEngine(IXLWorksheet ws)
        {
            _ws = ws;
            // parse multi-cell range references ($A2:B$4)
            IdentifierChars = "$:!";
        }

        public override object GetExternalObject(string identifier)
        {
            //if (!XLHelper.IsValidA1Address(identifier)) return null;
            //String wsName;
            if (identifier.Contains("!"))
            {
                var wsName = identifier.Substring(0, identifier.IndexOf("!"));
                return new CellRangeReference(_ws.Workbook.Worksheet(wsName).Range(identifier.Substring(identifier.IndexOf("!") + 1)), this);
            }
            return new CellRangeReference(_ws.Range(identifier), this);
        }


    }

    public class CellRangeReference : IValueObject, IEnumerable
    {
        private IXLRange _range;
        private XLCalcEngine _ce;
        public CellRangeReference(IXLRange range, XLCalcEngine ce)
        {
            _range = range;
            _ce = ce;
        }

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
                    return _ce.Evaluate(f);

            }
            finally
            {
                _evaluating = false;
            }
        }
    }
}
