using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLIgnoredErrors: IEnumerable<IXLIgnoredError>
    {
        void Clear();
        void Add(XLIgnoredErrorType type, IXLRange range);
    }
}
