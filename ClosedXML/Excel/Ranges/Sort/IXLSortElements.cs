using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLSortElements: IEnumerable<IXLSortElement>
    {
        void Add(Int32 elementNumber);
        void Add(Int32 elementNumber, XLSortOrder sortOrder);
        void Add(Int32 elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks);
        void Add(Int32 elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase);

        void Add(String elementNumber);
        void Add(String elementNumber, XLSortOrder sortOrder);
        void Add(String elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks);
        void Add(String elementNumber, XLSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase);

        void Clear();

        void Remove(Int32 elementNumber);
    }
}
