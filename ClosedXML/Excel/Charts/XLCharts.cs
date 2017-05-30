using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCharts: IXLCharts
    {
        private List<IXLChart> charts = new List<IXLChart>();
        public IEnumerator<IXLChart> GetEnumerator()
        {
            return charts.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IXLChart chart)
        {
            charts.Add(chart);
        }
    }
}
