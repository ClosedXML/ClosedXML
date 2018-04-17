using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLDataValidations: IEnumerable<IXLDataValidation>, IDisposable
    {
        void Add(IXLDataValidation dataValidation);
        Boolean ContainsSingle(IXLRange range);
        
    }
}
