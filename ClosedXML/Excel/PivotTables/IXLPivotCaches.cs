using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A collection of <see cref="IXLPivotCache">pivot caches</see>. Pivot cache
    /// can be added from a <see cref="IXLRange"/> or a <see cref="IXLTable"/>.
    /// </summary>
    public interface IXLPivotCaches : IEnumerable<IXLPivotCache>
    {
        /// <summary>
        /// Add a new pivot cache for the range. If the range area is same as
        /// an area of a table, the created cache will reference the table
        /// as source of data instead of a range of cells.
        /// </summary>
        /// <param name="range">Range for which to create the pivot cache.</param>
        /// <returns>The pivot cache for the range.</returns>
        IXLPivotCache Add(IXLRange range);
    }
}
