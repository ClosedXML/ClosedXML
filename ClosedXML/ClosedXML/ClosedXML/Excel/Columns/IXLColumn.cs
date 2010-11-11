using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLColumn : IXLRangeBase
    {
        Double Width { get; set; }
        void Delete();
        Int32 ColumnNumber();
        String ColumnLetter();
        void InsertColumnsAfter(Int32 numberOfColumns);
        void InsertColumnsBefore(Int32 numberOfColumns);
        void Clear();

        IXLCell Cell(int row);
        void AdjustToContents();
        void Hide();
        void Unhide();
        Boolean IsHidden { get; }
        Int32 OutlineLevel { get; set; }
        void Group(Boolean collapse = false);
        void Group(Int32 outlineLevel, Boolean collapse = false);
        void Ungroup(Boolean fromAll = false);
        void Collapse();
        void Expand();

    }
}
