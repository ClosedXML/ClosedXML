#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotValues : IEnumerable<IXLPivotValue>
    {
        /// <summary>
        /// Add a new value field to the pivot table. If addition would cause, the
        /// <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
        /// <see cref="IXLPivotTable.ColumnLabels"/>. The added field will use passed
        /// <paramref name="sourceName"/> as the <see cref="IXLPivotField.CustomName"/>.
        /// </summary>
        /// <param name="sourceName">The <see cref="IXLPivotField.SourceName"/> that is used as a
        ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
        /// <returns>Newly added field.</returns>
        IXLPivotValue Add(String sourceName);

        /// <summary>
        /// Add a new value field to the pivot table. If addition would cause, the
        /// <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
        /// <see cref="IXLPivotTable.ColumnLabels"/>.
        /// </summary>
        /// <param name="sourceName">The <see cref="IXLPivotField.SourceName"/> that is used as a
        ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
        /// <param name="customName">The added data field <see cref="IXLPivotField.CustomName"/>.</param>
        /// <returns>Newly added field.</returns>
        IXLPivotValue Add(String sourceName, String customName);

        void Clear();

        Boolean Contains(String customName);

        Boolean Contains(IXLPivotValue pivotValue);

        IXLPivotValue Get(String customName);

        IXLPivotValue Get(Int32 index);

        Int32 IndexOf(String customName);

        Int32 IndexOf(IXLPivotValue pivotValue);

        void Remove(String customName);
    }
}
