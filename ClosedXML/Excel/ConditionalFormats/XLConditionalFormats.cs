using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormats : IXLConditionalFormats
    {
        private readonly List<IXLConditionalFormat> _conditionalFormats = new List<IXLConditionalFormat>();
        public void Add(IXLConditionalFormat conditionalFormat)
        {
            _conditionalFormats.Add(conditionalFormat);
        }

        private bool IsRangeAbove(IXLRangeAddress newAddr, IXLRangeAddress addr)
        {
            return newAddr.FirstAddress.ColumnNumber == addr.FirstAddress.ColumnNumber
                   && newAddr.LastAddress.ColumnNumber == addr.LastAddress.ColumnNumber
                   && newAddr.FirstAddress.RowNumber < addr.FirstAddress.RowNumber
                   && (newAddr.LastAddress.RowNumber+1).Between(addr.FirstAddress.RowNumber, addr.LastAddress.RowNumber);
        }

        private bool IsRangeToLeft(IXLRangeAddress newAddr, IXLRangeAddress addr)
        {
            return newAddr.FirstAddress.RowNumber == addr.FirstAddress.RowNumber
                   && newAddr.LastAddress.RowNumber == addr.LastAddress.RowNumber
                   && newAddr.FirstAddress.ColumnNumber < addr.FirstAddress.ColumnNumber
                   && (newAddr.LastAddress.ColumnNumber+1).Between(addr.FirstAddress.ColumnNumber, addr.LastAddress.ColumnNumber);
        }

        public IEnumerator<IXLConditionalFormat> GetEnumerator()
        {
            return _conditionalFormats.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(Predicate<IXLConditionalFormat> predicate)
        {
            _conditionalFormats.Where(cf=>predicate(cf)).ForEach(cf=>cf.Range.Dispose());
            _conditionalFormats.RemoveAll(predicate);
        }

        /// <summary>
        /// The method consolidate the same conditional formats, which are located in adjacent ranges.
        /// </summary>
        internal void Consolidate()
        {
            var formats = _conditionalFormats
                .OrderByDescending(x => x.Range.RangeAddress.FirstAddress.RowNumber)
                .ThenByDescending(x => x.Range.RangeAddress.FirstAddress.ColumnNumber);

            var orderedFormats = formats.ToList();

            foreach (var item in formats)
            {
                var itemAddr = item.Range.RangeAddress;
                var itemRowNum = itemAddr.FirstAddress.RowNumber;

                Func<IXLConditionalFormat, bool> IsSameFormat = f => f != item && f.Range.Worksheet.Position == item.Range.Worksheet.Position &&
                                                             XLConditionalFormat.NoRangeComparer.Equals(f, item);

                // search for an adjacent range
                var format = orderedFormats
                    .TakeWhile(f => f.Range.RangeAddress.FirstAddress.RowNumber >= itemRowNum)
                    .FirstOrDefault(f => (IsRangeAbove(itemAddr, f.Range.RangeAddress) || IsRangeToLeft(itemAddr, f.Range.RangeAddress)) && IsSameFormat(f));
                if (format != null)
                {
                    Merge(format, item);
                    _conditionalFormats.Remove(item);
                    orderedFormats.Remove(item);
                    // compress with bottom range
                    var newaddr = format.Range.RangeAddress;
                    var newRowNum = newaddr.FirstAddress.RowNumber;
                    var bottom = orderedFormats
                        .TakeWhile(f => f.Range.RangeAddress.FirstAddress.RowNumber >= newRowNum)
                        .FirstOrDefault(f => IsRangeAbove(newaddr, f.Range.RangeAddress) && IsSameFormat(f));
                    if (bottom != null)
                    {
                        Merge(bottom, format);
                        _conditionalFormats.Remove(format);
                        orderedFormats.Remove(format);
                    }
                    continue;
                }

                // search for an encompassable range
                format = _conditionalFormats.FirstOrDefault(f => f.Range.Contains(item.Range) && IsSameFormat(f));
                if (format != null)
                {
                    _conditionalFormats.Remove(item);
                    orderedFormats.Remove(item);
                }
            }
        }

        private static void Merge(IXLConditionalFormat format, IXLConditionalFormat item)
        {
            foreach (var v in format.Values.ToList())
                format.Values[v.Key] = item.Values[v.Key];
            format.Range.RangeAddress.FirstAddress = item.Range.RangeAddress.FirstAddress;
        }

        public void RemoveAll()
        {
            _conditionalFormats.ForEach(cf => cf.Range.Dispose());
            _conditionalFormats.Clear();
        }
    }
}
