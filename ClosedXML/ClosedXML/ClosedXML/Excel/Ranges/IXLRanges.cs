using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRanges: IEnumerable<IXLRange>
    {
        /// <summary>
        /// Clears the contents of the ranges (including styles).
        /// </summary>
        void Clear();
        /// <summary>
        /// Adds the specified range to this group.
        /// </summary>
        /// <param name="range">The range to add to this group.</param>
        void Add(IXLRange range);
        /// <summary>
        /// Removes the specified range from this group.
        /// </summary>
        /// <param name="range">The range to remove from this group.</param>
        void Remove(IXLRange range);

        Boolean Contains(IXLRange range);

        IXLStyle Style { get; set; }

        IXLDataValidation DataValidation { get; }
    }
}
