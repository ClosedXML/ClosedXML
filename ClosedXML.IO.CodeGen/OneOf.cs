using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.IO.CodeGen;

public readonly record struct OneOf<T1, T2>
{
    private readonly T1? _t1;
    private readonly T2? _t2;

    private OneOf(T1? t1, T2? t2)
    {
        _t1 = t1;
        _t2 = t2;
    }

    public static implicit operator OneOf<T1, T2>(T1 value) => new(value, default);
    public static implicit operator OneOf<T1, T2>(T2 value) => new(default, value);

    public bool TryPickT1([NotNullWhen(true)] out T1? t1)
    {
        t1 = _t1!;
        return !EqualityComparer<T1?>.Default.Equals(_t1, default);
    }

    internal bool TryPickT2([NotNullWhen(true)] out T2? t2)
    {
        t2 = _t2!;
        return !EqualityComparer<T2?>.Default.Equals(_t2, default);
    }
}
