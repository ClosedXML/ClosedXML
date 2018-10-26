using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLDataValidations : IEnumerable<IXLDataValidation>
    {
        void Add(IXLDataValidation dataValidation);

        Boolean ContainsSingle(IXLRange range);

        void Delete(Predicate<IXLDataValidation> predicate);
    }
}
