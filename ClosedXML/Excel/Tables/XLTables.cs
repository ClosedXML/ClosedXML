using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    public class XLTables : IXLTables
    {
        private readonly Dictionary<String, IXLTable> _tables = new Dictionary<String, IXLTable>();

        #region IXLTables Members

        public IEnumerator<IXLTable> GetEnumerator()
        {
            return _tables.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IXLTable table)
        {
            _tables.Add(table.Name, table);
        }

        public IXLTable Table(Int32 index)
        {
            return _tables.ElementAt(index).Value;
        }

        public IXLTable Table(String name)
        {
            return _tables[name];
        }

        #endregion

        public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            _tables.Values.ForEach(t => t.Clear(clearOptions));
            return this;
        }

        public void Remove(Int32 index)
        {
            _tables.Remove(_tables.ElementAt(index).Key);
        }
        public void Remove(String name)
        {
            _tables.Remove(name);
        }
    }
}