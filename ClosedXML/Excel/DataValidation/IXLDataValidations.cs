using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLDataValidations : IEnumerable<IXLDataValidation>
    {
        IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Add data validation rule to the collection. If the specified rule refers to another
        /// worksheet than the collection, the copy will be created and its ranges will refer
        /// to the worksheet of the collection. Otherwise the original instance will be placed
        /// in the collection.
        /// </summary>
        /// <param name="dataValidation">A data validation rule to add.</param>
        /// <returns>The instance that has actually been added in the collection
        /// (may be a copy of the specified one).</returns>
        IXLDataValidation Add(IXLDataValidation dataValidation);

        Boolean ContainsSingle(IXLRange range);

        void Delete(Predicate<IXLDataValidation> predicate);
    }
}
