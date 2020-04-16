using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLHyperlinks: IEnumerable<XLHyperlink>
    {
        void Add(XLHyperlink hyperlink);
        void Delete(XLHyperlink hyperlink);
        void Delete(IXLAddress address);
        bool TryDelete(IXLAddress address);
        XLHyperlink Get(IXLAddress address);
        bool TryGet(IXLAddress address, out XLHyperlink hyperlink);
    }
}
