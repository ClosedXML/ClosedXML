using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLTables: IEnumerable<IXLTable>
    {
        void Add(IXLTable table);
        IXLTable Table(Int32 index);
        IXLTable Table(String name);
    }
}
