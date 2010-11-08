using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRow : IXLRangeBase
    {
        Double Height { get; set; }
        void Delete();
        Int32 RowNumber();
        void InsertRowsBelow(Int32 numberOfRows);
        void InsertRowsAbove(Int32 numberOfRows);
        void Clear();

        IXLCell Cell(Int32 column);
        IXLCell Cell(String column);

        void AdjustToContents();
    }
}
