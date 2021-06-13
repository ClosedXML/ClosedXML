// Keep this file CodeMaid organised and cleaned
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLRangeColumnComparer : IComparer<IXLRangeColumn>
    {
        private readonly IXLSortElements _sortElements;

        internal XLRangeColumnComparer(IXLSortElements sortElements)
        {
            this._sortElements = sortElements;
        }

        public int Compare(IXLRangeColumn x, IXLRangeColumn y)
        {
            foreach (IXLSortElement sortElement in _sortElements)
            {
                var comparison = ((XLSortElement)sortElement).CellComparer.Compare((XLCell)x.Cell(sortElement.ElementNumber), (XLCell)y.Cell(sortElement.ElementNumber));
                if (comparison != 0)
                    return comparison;
            }

            return 0;
        }
    }
}
