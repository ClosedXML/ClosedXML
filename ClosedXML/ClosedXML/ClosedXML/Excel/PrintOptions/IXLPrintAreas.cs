using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPrintAreas: IEnumerable<IXLRange>
    {
        void Clear();
        void Add(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn);
        void Add(String rangeAddress);
        void Add(String firstCellAddress, String lastCellAddress);
        void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress);
    }
}
