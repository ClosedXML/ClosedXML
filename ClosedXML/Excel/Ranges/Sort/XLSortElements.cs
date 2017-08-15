using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLSortElements: IXLSortElements
    {
        List<IXLSortElement> elements = new List<IXLSortElement>();
        public void Add(Int32 elementNumber)
        {
            Add(elementNumber, XLSortOrder.Ascending);
        }
        public void Add(Int32 elementNumber, XLSortOrder sortOrder)
        {
            Add(elementNumber, sortOrder, true);
        }
        public void Add(Int32 elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks)
        {
            Add(elementNumber, sortOrder, ignoreBlanks, false);
        }
        public void Add(Int32 elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase)
        {
            elements.Add(new XLSortElement()
            {
                ElementNumber = elementNumber,
                SortOrder = sortOrder,
                IgnoreBlanks = ignoreBlanks,
                MatchCase = matchCase
            });
        }

        public void Add(String elementNumber)
        {
            Add(elementNumber, XLSortOrder.Ascending);
        }
        public void Add(String elementNumber, XLSortOrder sortOrder)
        {
            Add(elementNumber, sortOrder, true);
        }
        public void Add(String elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks)
        {
            Add(elementNumber, sortOrder, ignoreBlanks, false);
        }
        public void Add(String elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase)
        {
            elements.Add(new XLSortElement()
            {
                ElementNumber = XLHelper.GetColumnNumberFromLetter(elementNumber),
                SortOrder = sortOrder,
                IgnoreBlanks = ignoreBlanks,
                MatchCase = matchCase
            });
        }

        public IEnumerator<IXLSortElement> GetEnumerator()
        {
            return elements.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Clear()
        {
            elements.Clear();
        }

        public void Remove(Int32 elementNumber)
        {
            elements.RemoveAt(elementNumber - 1);
        }
    }
}
