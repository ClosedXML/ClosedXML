using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLTables: IXLTables
    {
        private List<IXLTable> tables = new List<IXLTable>();
        public IEnumerator<IXLTable> GetEnumerator()
        {
            return tables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IXLTable table)
        {
            tables.Add(table);
        }

        public IXLTable Table(Int32 index)
        {
            return tables[index];
        }

        public IXLTable Table(String name)
        {
            return tables.Where(t => t.Name == name).Single();
        }
    }
}
