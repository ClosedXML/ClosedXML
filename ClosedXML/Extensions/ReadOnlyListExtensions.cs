#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Extensions
{
    internal static class ReadOnlyListExtensions
    {
        /// <summary>
        /// Searches for the specified item and returns the zero-based index of the first occurrence.
        /// </summary>
        /// <returns>Index of found item, -1 if item is not found.</returns>
        public static Int32 IndexOf<T>(this IReadOnlyList<T> source, T item, IEqualityComparer<T> comparer)
        {
            for (var i = 0; i < source.Count; ++i)
            {
                if (comparer.Equals(source[i], item))
                    return i;
            }

            return -1;
        }
    }
}
