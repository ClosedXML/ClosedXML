using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPrintAreas : IXLPrintAreas
    {
        List<IXLRange> ranges = new List<IXLRange>();
        private XLWorksheet worksheet;
        public XLPrintAreas(XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet worksheet)
        {
            ranges = defaultPrintAreas.ranges.ToList();
            this.worksheet = worksheet;
        }

        public void Clear()
        {
            ranges.Clear();
        }

        public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            ranges.Add(worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
        }

        public void Add(string rangeAddress)
        {
            ranges.Add(worksheet.Range(rangeAddress));
        }

        public void Add(string firstCellAddress, string lastCellAddress)
        {
            ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            return ranges.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
