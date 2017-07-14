using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLHyperlinks: IEnumerable<XLHyperlink>
    {
        void Add(XLHyperlink hyperlink);
        void Delete(XLHyperlink hyperlink);
        void Delete(IXLAddress address);
        
    }
}
