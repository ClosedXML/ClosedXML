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
            _conditionalFormats.Where(cf => predicate(cf)).SelectMany(cf => cf.Ranges).ForEach(range => range.Dispose());
            _conditionalFormats.RemoveAll(predicate);
        }

        /// <summary>
        /// The method consolidate the same conditional formats, which are located in adjacent ranges.
        /// </summary>
        internal void Consolidate()
        {
            var formats = _conditionalFormats
                .Where(cf => cf.Ranges.Any())
                .ToList();
            _conditionalFormats.Clear();

            while (formats.Count > 0)
            {
                var item = formats.First();

                var ranges = item.Ranges;
                Func<IXLConditionalFormat, bool> IsSameFormat = f =>
                    f != item && f.Ranges.First().Worksheet.Position == item.Ranges.First().Worksheet.Position &&
                    XLConditionalFormat.NoRangeComparer.Equals(f, item);
                int i = 1;
                bool stop = false;
                List<IXLConditionalFormat> similarFormats = new List<IXLConditionalFormat>();
                do
                {
                    stop = (i >= formats.Count);

                    if (!stop)
                    {
                        var nextFormat = formats[i];
                        if (IsSameFormat(nextFormat))
                        {
                            similarFormats.Add(nextFormat);
                            nextFormat.Ranges.ForEach(r => ranges.Add(r));
                        }
                        else if (ranges.Any(left => nextFormat.Ranges.Any(right => left.Intersects(right))))
                        {
                            // if we reached the rule intersecting any of captured ranges stop for not breaking the priorities
                            stop = true;
                        }
                    }

                    i++;
                } while (!stop);

                var baseCell = item.Ranges.First().FirstCell() as XLCell;

                var consRanges = ranges.Consolidate();
                item.Ranges.RemoveAll();
                consRanges.ForEach(r => item.Ranges.Add(r));

                var targetCell = item.Ranges.First().FirstCell() as XLCell;
                (item as XLConditionalFormat).AdjustFormulas(baseCell, targetCell);

                _conditionalFormats.Add(item);

                similarFormats.ForEach(cf => formats.Remove(cf));
                formats.Remove(item);
            }
        }

        public void RemoveAll()
        {
            _conditionalFormats.SelectMany(cf => cf.Ranges).ForEach(range => range.Dispose());
            _conditionalFormats.Clear();
        }

        /// <summary>
        /// Reorders the according to original priority. Done during load process
        /// </summary>
        public void ReorderAccordingToOriginalPriority()
        {
            var reorderedFormats = _conditionalFormats.OrderBy(cf => (cf as XLConditionalFormat).OriginalPriority).ToList();
            _conditionalFormats.Clear();
            _conditionalFormats.AddRange(reorderedFormats);
        }
    }
}
