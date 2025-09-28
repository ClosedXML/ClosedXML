using System.Collections.Generic;

namespace ClosedXML.Utils;

internal interface IReadOnlyBiDictionary<TKey, TValue> : IReadOnlyDictionary<TKey, TValue>
    where TKey : notnull
{
    /// <summary>
    /// Return the key associated with the value. The bi-dictionary does allow duplicate values.
    /// In case of duplicates, the earliest added entry will be returned.
    /// </summary>
    TKey this[TValue value] { get; }

    /// <summary>
    /// Does the dictionary contain the value?
    /// </summary>
    bool ContainsValue(TValue value);
}
