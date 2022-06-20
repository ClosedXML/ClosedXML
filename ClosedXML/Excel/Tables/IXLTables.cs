// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLTables : IEnumerable<IXLTable>
    {
        void Add(IXLTable table);

        /// <summary>
        /// Clears the contents of these tables.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All);

        bool Contains(string name);

        void Remove(int index);

        void Remove(string name);

        IXLTable Table(int index);

        IXLTable Table(string name);

        bool TryGetTable(string tableName, out IXLTable table);
    }
}
