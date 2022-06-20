using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLSortElements: IEnumerable<IXLSortElement>
    {
        void Add(int elementNumber);
        void Add(int elementNumber, XLSortOrder sortOrder);
        void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks);
        void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

        void Add(string elementNumber);
        void Add(string elementNumber, XLSortOrder sortOrder);
        void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks);
        void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

        void Clear();

        void Remove(int elementNumber);
    }
}
