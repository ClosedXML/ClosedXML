using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLConditionalFormats: IEnumerable<IXLConditionalFormat>
    {
        void Add(IXLConditionalFormat conditionalFormat);
        void RemoveAll();
        void Remove(Predicate<IXLConditionalFormat> predicate);
    }
}
