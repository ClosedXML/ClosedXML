using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    /// <summary>
    /// Base interface for an abstract repository.
    /// </summary>
    internal interface IXLRepository
    {
        /// <summary>
        /// Clear the repository;
        /// </summary>
        void Clear();
    }

    internal interface IXLRepository<Tkey, Tvalue> : IXLRepository, IEnumerable<Tvalue>
        where Tkey : struct, IEquatable<Tkey>
        where Tvalue : class
    {
        /// <summary>
        /// Put the <paramref name="value"/> into the repository under the specified <paramref name="key"/>
        /// if there is no such key present.
        /// </summary>
        /// <param name="key">Key to identify the value.</param>
        /// <param name="value">Value to put into the repository if key does not exist.</param>
        /// <returns>Value stored in the repository under the specified <paramref name="key"/>. If key already existed
        /// returned value may differ from the input one.</returns>
        Tvalue Store(ref Tkey key, Tvalue value);
    }
}
