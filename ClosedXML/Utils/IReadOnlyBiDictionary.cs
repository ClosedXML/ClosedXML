using System.Collections.Generic;

namespace ClosedXML.Utils;

internal interface IReadOnlyBiDictionary<TKey, TValue> : IReadOnlyCollection<KeyValuePair<TKey, TValue>>
    where TKey : notnull
{
    TValue this[TKey key] { get; }
}
