using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel
{
    public interface IXLDefinedNames : IEnumerable<IXLDefinedName>
    {
        /// <inheritdoc cref="DefinedName"/>
        [Obsolete($"Use {nameof(DefinedName)} instead.")]
        IXLDefinedName NamedRange(String name);

        /// <summary>
        /// Gets the specified defined name.
        /// </summary>
        /// <param name="name">Name identifier.</param>
        /// <exception cref="KeyNotFoundException">Name wasn't found.</exception>
        IXLDefinedName DefinedName(String name);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <exception cref="ArgumentException">The name or address is invalid.</exception>
        IXLDefinedName Add(String name, String rangeAddress);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="range">The range to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXLDefinedName Add(String name, IXLRange range);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXLDefinedName Add(String name, IXLRanges ranges);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name or address is invalid.</exception>
        IXLDefinedName Add(String name, String rangeAddress, String? comment);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="range">The range to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXLDefinedName Add(String name, IXLRange range, String? comment);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXLDefinedName Add(String name, IXLRanges ranges, String? comment);

        /// <summary>
        /// Deletes the specified defined name.  Deleting defined name doesn't delete referenced
        /// cells.
        /// </summary>
        /// <param name="name">Name identifier to delete.</param>
        void Delete(String name);

        /// <summary>
        /// Deletes the specified defined name's index. Deleting defined name doesn't delete
        /// referenced cells.
        /// </summary>
        /// <param name="index">Index of the defined name to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">The index is outside of named ranges array.</exception>
        void Delete(Int32 index);

        /// <summary>
        /// Deletes all defined names of this collection, i.e. a workbook or a sheet. Deleting
        /// defined name doesn't delete referenced cells.
        /// </summary>
        void DeleteAll();

        Boolean TryGetValue(String name, [NotNullWhen(true)] out IXLDefinedName? range);

        Boolean Contains(String name);

        /// <summary>
        /// Returns a subset of defined names that do not have invalid references.
        /// </summary>
        IEnumerable<IXLDefinedName> ValidNamedRanges();

        /// <summary>
        /// Returns a subset of defined names that do have invalid references.
        /// </summary>
        IEnumerable<IXLDefinedName> InvalidNamedRanges();
    }
}
