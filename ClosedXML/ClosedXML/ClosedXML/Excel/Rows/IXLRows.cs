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
    }
}
