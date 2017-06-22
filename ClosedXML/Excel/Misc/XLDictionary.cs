using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLDictionary<T>: Dictionary<Int32, T>
    {
        public XLDictionary()
        {
            
        }
        public XLDictionary(XLDictionary<T> other)
        {
            other.Values.ForEach(Add);
        }

        public void Initialize(T value)
        {
            if (Count > 0)
                Clear();

            Add(value);
        }

        public void Add(T value)
        {
            Add(Count + 1, value);
        }
    }
}
