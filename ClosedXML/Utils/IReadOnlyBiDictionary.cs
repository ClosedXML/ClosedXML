using System;
using System.Collections.Generic;

namespace ClosedXML.Utils;

internal interface IReadOnlyBiDictionary<TKey, TValue> : IReadOnlyCollection<KeyValuePair<TKey, TValue>>
    where TKey : notnull
    where TValue : IEquatable<TValue>
{
    TValue this[TKey key] { get; }
}
