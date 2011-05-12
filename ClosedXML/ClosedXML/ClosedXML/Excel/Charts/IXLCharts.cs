using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCharts: IEnumerable<IXLChart>
    {
        void Add(IXLChart chart);
    }
}
