// Keep this file CodeMaid organised and cleaned
using System;
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
                        switch (thisCell.DataType)
                        {
                            case XLDataType.Text:
                                comparison = e.MatchCase
                                                 ? thisCell.InnerText.CompareTo(otherCell.InnerText)
                                                 : String.Compare(thisCell.InnerText, otherCell.InnerText, true);
                                break;

                            case XLDataType.TimeSpan:
                                comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                                break;

                            case XLDataType.DateTime:
                                comparison = thisCell.GetDateTime().CompareTo(otherCell.GetDateTime());
                                break;

                            case XLDataType.Number:
                                comparison = thisCell.GetDouble().CompareTo(otherCell.GetDouble());
                                break;

                            case XLDataType.Boolean:
                                comparison = thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
                                break;

                            default:
                                throw new NotImplementedException();
                        }
                    }
                    else if (e.MatchCase)
                        comparison = String.Compare(thisCell.GetString(), otherCell.GetString(), true);
                    else
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }

                if (comparison != 0)
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : -comparison;
            }

            return 0;
        }
    }
}
