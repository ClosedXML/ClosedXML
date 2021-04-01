// Keep this file CodeMaid organised and cleaned
using System;
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
            foreach (IXLSortElement e in _sortElements)
            {
                var thisCell = (XLCell)x.Cell(e.ElementNumber);
                var otherCell = (XLCell)y.Cell(e.ElementNumber);
                int comparison;
                bool thisCellIsBlank = thisCell.IsEmpty();
                bool otherCellIsBlank = otherCell.IsEmpty();
                if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
                {
                    if (thisCellIsBlank && otherCellIsBlank)
                        comparison = 0;
                    else
                    {
                        if (thisCellIsBlank)
                            comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                        else
                            comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                    }
                }
                else
                {
                    if (thisCell.DataType == otherCell.DataType)
                    {
                        if (thisCell.DataType == XLDataType.Text)
                        {
                            comparison = e.MatchCase
                                             ? thisCell.InnerText.CompareTo(otherCell.InnerText)
                                             : String.Compare(thisCell.InnerText, otherCell.InnerText, true);
                        }
                        else if (thisCell.DataType == XLDataType.TimeSpan)
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                        else
                            comparison = Double.Parse(thisCell.InnerText, XLHelper.NumberStyle, XLHelper.ParseCulture).CompareTo(Double.Parse(otherCell.InnerText, XLHelper.NumberStyle, XLHelper.ParseCulture));
                    }
                    else if (e.MatchCase)
                        comparison = String.Compare(thisCell.GetString(), otherCell.GetString(), true);
                    else
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }

                if (comparison != 0)
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
            }

            return 0;
        }
    }
}
