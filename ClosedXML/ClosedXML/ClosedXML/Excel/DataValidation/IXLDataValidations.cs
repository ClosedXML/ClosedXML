using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDataValidations: IEnumerable<IXLDataValidation>
    {
        void Add(IXLDataValidation dataValidation);
        Boolean ContainsSingle(IXLRange range);
        
    }
}
