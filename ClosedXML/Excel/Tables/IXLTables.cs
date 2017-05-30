using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLTables: IEnumerable<IXLTable>
    {
        void Add(IXLTable table);
        IXLTable Table(Int32 index);
        IXLTable Table(String name);

        /// <summary>
        /// Clears the contents of these tables.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        void Remove(Int32 index);
        void Remove(String name);
    }
}
