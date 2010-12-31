using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLSheetView: IXLSheetView
    {
        public Int32 SplitRow { get; set; }
        public Int32 SplitColumn { get; set; }
        public Boolean FreezePanes { get; set; }
        public void FreezeRows(Int32 rows)
        {
            SplitRow = rows;
            FreezePanes = true;
        }
        public void FreezeColumns(Int32 columns)
        {
            SplitColumn = columns;
            FreezePanes = true;
        }
        public void Freeze(Int32 rows, Int32 columns)
        {
            SplitRow = rows;
            SplitColumn = columns;
            FreezePanes = true;
        }
    }
}
