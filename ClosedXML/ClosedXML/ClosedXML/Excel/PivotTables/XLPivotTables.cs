using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotTables: IXLPivotTables
    {
        private readonly Dictionary<String, XLPivotTable> _pivotTables = new Dictionary<string, XLPivotTable>();
        public IEnumerator<IXLPivotTable> GetEnumerator()
        {
            return _pivotTables.Values.Cast<IXLPivotTable>().GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public XLPivotTable PivotTable(String name)
        {
            return _pivotTables[name];
        }
        IXLPivotTable IXLPivotTables.PivotTable(String name)
        {
            return PivotTable(name);
        }
        public void Delete(String name)
        {
            _pivotTables.Remove(name);
        }
        public void DeleteAll()
        {
            _pivotTables.Clear();
        }

        public void Add(String name, XLPivotTable pivotTable)
        {
            _pivotTables.Add(name, pivotTable);
        }
    }
}
