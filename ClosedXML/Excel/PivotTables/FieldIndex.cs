using System;

namespace ClosedXML.Excel;

/// <summary>
/// A type for field index, so there is a better idea what is a semantic content of some
/// variable/props. Not detrimental to performance, JIT will inline struct to int.
/// </summary>
internal readonly record struct FieldIndex
{
    internal FieldIndex(int value)
    {
        if (value < 0 && value != -2)
            throw new ArgumentOutOfRangeException();

        Value = value;
    }

    /// <summary>
    /// Index of a field in <see cref="XLPivotTable.PivotFields"/>. Can be -2 for 'data' field.
    /// </summary>
    public int Value { get; }

    public static implicit operator int(FieldIndex index) => index.Value;

    public static implicit operator FieldIndex(int index) => new(index);
}
