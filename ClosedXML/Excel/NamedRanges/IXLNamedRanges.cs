using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLNamedRanges: IEnumerable<IXLNamedRange>
    {
        /// <summary>
        /// Gets the specified named range.
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        IXLNamedRange NamedRange(string rangeName);
        
        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <returns></returns>
        IXLNamedRange Add(string rangeName, string rangeAddress);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="range">The range to add.</param>
        /// <returns></returns>
        IXLNamedRange Add(string rangeName, IXLRange range);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <returns></returns>
        IXLNamedRange Add(string rangeName, IXLRanges ranges);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        IXLNamedRange Add(string rangeName, string rangeAddress, string comment);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="range">The range to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        IXLNamedRange Add(string rangeName, IXLRange range, string comment);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        IXLNamedRange Add(string rangeName, IXLRanges ranges, string comment);

        /// <summary>
        /// Deletes the specified named range (not the cells).
        /// </summary>
        /// <param name="rangeName">Name of the range to delete.</param>
        void Delete(string rangeName);

        /// <summary>
        /// Deletes the specified named range's index (not the cells).
        /// </summary>
        /// <param name="rangeIndex">Index of the named range to delete.</param>
        void Delete(int rangeIndex);


        /// <summary>
        /// Deletes all named ranges (not the cells).
        /// </summary>
        void DeleteAll();

        bool TryGetValue(string name, out IXLNamedRange range);

        bool Contains(string name);

        /// <summary>
        /// Returns a subset of named ranges that do not have invalid references.
        /// </summary>
        IEnumerable<IXLNamedRange> ValidNamedRanges();

        /// <summary>
        /// Returns a subset of named ranges that do have invalid references.
        /// </summary>
        IEnumerable<IXLNamedRange> InvalidNamedRanges();
    }
}
