// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLRangeRowComparer : IComparer<IXLRangeRow>
    {
        private readonly IXLSortElements _sortElements;

        internal XLRangeRowComparer(IXLSortElements sortElements)
        {
            this._sortElements = sortElements;
        }

        public int Compare(IXLRangeRow x, IXLRangeRow y)
        {
            foreach (var sortElement in _sortElements)
            {
                var comparison = ((XLSortElement)sortElement).CellComparer.Compare((XLCell)x.Cell(sortElement.ElementNumber), (XLCell)y.Cell(sortElement.ElementNumber));
                if (comparison != 0)
                    return comparison;
            }

            return 0;
        }
    }
}
