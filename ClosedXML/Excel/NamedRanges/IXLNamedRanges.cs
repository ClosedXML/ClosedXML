using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel
{
    public interface IXLNamedRanges : IEnumerable<IXLNamedRange>
    {
        /// <summary>
        /// Gets the specified named range.
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        /// <exception cref="ArgumentException">Range wasn't found.</exception>
        IXLNamedRange NamedRange(String rangeName);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <exception cref="ArgumentException">The name or address is invalid.</exception>
        IXLNamedRange Add(String rangeName, String rangeAddress);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="range">The range to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXLNamedRange Add(String rangeName, IXLRange range);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the range to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXLNamedRange Add(String rangeName, IXLRanges ranges);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name or address is invalid.</exception>
        IXLNamedRange Add(String rangeName, String rangeAddress, String? comment);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="range">The range to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXLNamedRange Add(String rangeName, IXLRange range, String? comment);

        /// <summary>
        /// Adds a new named range.
        /// </summary>
        /// <param name="rangeName">Name of the ranges to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXLNamedRange Add(String rangeName, IXLRanges ranges, String? comment);

        /// <summary>
        /// Deletes the specified named range (not the cells).
        /// </summary>
        /// <param name="rangeName">Name of the range to delete.</param>
        void Delete(String rangeName);

        /// <summary>
        /// Deletes the specified named range's index (not the cells).
        /// </summary>
        /// <param name="rangeIndex">Index of the named range to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">The index is outside of named ranges array.</exception>
        void Delete(Int32 rangeIndex);

        /// <summary>
        /// Deletes all named ranges (not the cells).
        /// </summary>
        void DeleteAll();

        Boolean TryGetValue(String name, [NotNullWhen(true)] out IXLNamedRange? range);

        Boolean Contains(String name);

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
