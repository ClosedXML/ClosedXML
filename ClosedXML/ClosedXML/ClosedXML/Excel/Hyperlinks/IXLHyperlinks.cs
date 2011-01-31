using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLHyperlinks: IEnumerable<XLHyperlink>
    {
        void Add(XLHyperlink hyperlink);
        void Delete(XLHyperlink hyperlink);
        void Delete(IXLAddress address);
    }
}
