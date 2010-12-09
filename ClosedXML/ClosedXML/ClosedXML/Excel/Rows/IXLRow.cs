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
        void Hide();
        void Unhide();
        Boolean IsHidden { get; }
        Int32 OutlineLevel { get; set; }
        void Group();
        void Group(Boolean collapse);
        void Group(Int32 outlineLevel);
        void Group(Int32 outlineLevel, Boolean collapse);
        void Ungroup();
        void Ungroup(Boolean fromAll);
        void Collapse();
        void Expand();
    }
}
