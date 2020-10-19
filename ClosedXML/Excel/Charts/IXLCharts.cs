using System.Collections.Generic;
using System;

namespace ClosedXML.Excel
{
    public interface IXLCharts: IEnumerable<IXLChart>
    {
        IXLChart Add(IXLChart chart);

        IXLChart Chart(Int32 index);
    }
}
