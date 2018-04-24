using System;
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

        Boolean Contains(String name);

        void Remove(Int32 index);

        void Remove(String name);

        IXLTable Table(Int32 index);

        IXLTable Table(String name);
    }
}
