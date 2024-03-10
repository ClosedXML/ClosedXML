using System;
using System.Diagnostics;

namespace ClosedXML.Excel;

/// <summary>
/// A type for field index, so there is a better idea what is a semantic content of some
/// variable/props. Not detrimental to performance, JIT will inline struct to int.
/// </summary>
[DebuggerDisplay("{Value}")]
internal readonly record struct FieldIndex
{
    internal FieldIndex(int value)
    {
        if (value < 0 && value != -2)
            throw new ArgumentOutOfRangeException();

        Value = value;
    }

    /// <summary>
    /// The index of a 'data' field (<see cref="XLConstants.PivotTable.ValuesSentinalLabel"/>).
    /// </summary>
    internal static FieldIndex DataField => -2;

    /// <summary>
    /// Index of a field in <see cref="XLPivotTable.PivotFields"/>. Can be -2 for 'data' field,
    /// otherwise non-negative.
    /// </summary>
    internal int Value { get; }

    /// <summary>
    /// Is this index for a 'data' field?
    /// </summary>
    internal bool IsDataField => Value == -2;

    public static implicit operator int(FieldIndex index) => index.Value;

    public static implicit operator FieldIndex(int index) => new(index);
}
