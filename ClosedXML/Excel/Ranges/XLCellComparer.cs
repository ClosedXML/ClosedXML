// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCellComparer : IComparer<XLCell>
    {
        internal XLCellComparer(IXLSortElement sortElement)
            : this(sortElement.SortOrder, sortElement.MatchCase, sortElement.IgnoreBlanks)
        { }

        internal XLCellComparer(XLSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
        {
            this.SortOrder = sortOrder;
            this.MatchCase = matchCase;
            this.IgnoreBlanks = ignoreBlanks;
        }

        public bool IgnoreBlanks { get; }
        public bool MatchCase { get; }
        public XLSortOrder SortOrder { get; }

        public int Compare(XLCell x, XLCell y)
        {
            int comparison;
            bool thisCellIsBlank = x.IsEmpty();
            bool otherCellIsBlank = y.IsEmpty();
            if (IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
            {
                if (thisCellIsBlank && otherCellIsBlank)
                    comparison = 0;
                else
                {
                    if (thisCellIsBlank)
                        comparison = SortOrder == XLSortOrder.Ascending ? 1 : -1;
                    else
                        comparison = SortOrder == XLSortOrder.Ascending ? -1 : 1;
                }
            }
            else
            {
                if (x.DataType == y.DataType)
                {
                    comparison = x.DataType switch
                    {
                        XLDataType.Text => MatchCase
                                            ? x.InnerText.CompareTo(y.InnerText)
                                            : String.Compare(x.InnerText, y.InnerText, true),
                        XLDataType.TimeSpan => x.GetTimeSpan().CompareTo(y.GetTimeSpan()),
                        XLDataType.DateTime => x.GetDateTime().CompareTo(y.GetDateTime()),
                        XLDataType.Number => x.GetDouble().CompareTo(y.GetDouble()),
                        XLDataType.Boolean => x.GetBoolean().CompareTo(y.GetBoolean()),
                        _ => throw new NotImplementedException(),
                    };
                }
                else if (MatchCase)
                    comparison = String.Compare(x.GetString(), y.GetString(), true);
                else
                    comparison = x.GetString().CompareTo(y.GetString());
            }

            if (comparison != 0)
                return SortOrder == XLSortOrder.Ascending ? comparison : -comparison;

            return 0;
        }
    }
}
