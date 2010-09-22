using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLColumns: IEnumerable<IXLColumn>, IXLStylized
    {
        Double Width { get; set; }
        void Delete();
    }
}
