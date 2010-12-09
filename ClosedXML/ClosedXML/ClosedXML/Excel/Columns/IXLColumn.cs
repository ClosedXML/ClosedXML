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
