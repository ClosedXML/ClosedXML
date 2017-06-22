using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLCharts: IEnumerable<IXLChart>
    {
        void Add(IXLChart chart);
    }
}
