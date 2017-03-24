using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormats: IXLConditionalFormats
    {
        private readonly List<IXLConditionalFormat> _conditionalFormats = new List<IXLConditionalFormat>();
        public void Add(IXLConditionalFormat conditionalFormat)
        {
            _conditionalFormats.Add(conditionalFormat);
        }

        private bool RangeAbove(IXLRangeAddress newAddr, IXLRangeAddress addr)
        {
            return newAddr.FirstAddress.ColumnNumber == addr.FirstAddress.ColumnNumber
                   && newAddr.LastAddress.ColumnNumber == addr.LastAddress.ColumnNumber
                   && newAddr.FirstAddress.RowNumber < addr.FirstAddress.RowNumber
                   && (newAddr.LastAddress.RowNumber+1).Between(addr.FirstAddress.RowNumber, addr.LastAddress.RowNumber);
        }

        private bool RangeBefore(IXLRangeAddress newAddr, IXLRangeAddress addr)
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

        public void Compress()
        {
            var formats = _conditionalFormats
                .OrderByDescending(x=>x.Range.RangeAddress.FirstAddress.RowNumber)
                .ThenByDescending(x=>x.Range.RangeAddress.FirstAddress.ColumnNumber)
                .ToList();
            foreach (var item in formats)
            {
                var addr = item.Range.RangeAddress;
                var sameFormats = _conditionalFormats.Where(
                        f => f != item
                             && XLConditionalFormat.NoRangeComparer.Equals(f, item)
                             && f.Range.Worksheet.Position == item.Range.Worksheet.Position)
                    .ToArray();

                var format = sameFormats.FirstOrDefault(f => f.Range.Contains(item.Range));
                if (format != null)
                {
                    _conditionalFormats.Remove(item);
                    continue;
                }

                format = sameFormats.FirstOrDefault(f => RangeAbove(addr, f.Range.RangeAddress) || RangeBefore(addr, f.Range.RangeAddress));
                if (format != null)
                {
                    foreach (var v in format.Values.ToList())
                        format.Values[v.Key] = item.Values[v.Key];
                    format.Range.RangeAddress.FirstAddress = addr.FirstAddress;
                    _conditionalFormats.Remove(item);
                    continue;
                }
            }
        }

        public void RemoveAll()
        {
            _conditionalFormats.ForEach(cf => cf.Range.Dispose());
            _conditionalFormats.Clear();
        }
    }
}
