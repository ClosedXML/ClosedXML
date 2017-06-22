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
            _conditionalFormats.RemoveAll(predicate);
        }

        public void RemoveAll()
        {
            _conditionalFormats.Clear();
        }
    }
}
