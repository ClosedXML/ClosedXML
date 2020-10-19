using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public class XLCharts : IXLCharts
    {
        private List<IXLChart> charts = new List<IXLChart>();
        private readonly IXLWorksheet _worksheet;

        public XLCharts(IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public IEnumerator<IXLChart> GetEnumerator()
        {
            return charts.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLChart Chart(int index)
        {
            return charts[index];
        }

        public IXLChart Add(IXLChart chart)
        {
            charts.Add(chart);
            return chart;
        }
    }
}
