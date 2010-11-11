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
        void Group(Boolean collapse = false);
        void Group(Int32 outlineLevel, Boolean collapse = false);
        void Ungroup(Boolean fromAll = false);
        void Collapse();
        void Expand();
    }
}
