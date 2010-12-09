using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRows: IEnumerable<IXLRow>, IXLStylized
    {
        Double Height { set; }
        void Delete();
        void AdjustToContents();
        void Hide();
        void Unhide();
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
