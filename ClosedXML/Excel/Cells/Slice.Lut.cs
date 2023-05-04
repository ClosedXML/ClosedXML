#nullable disable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ClosedXML.Excel
{
    internal partial class Slice<TElement>
    {
        /// <summary>
        /// <para>
        /// Memory efficient look up table. The table is 2-level structure,
        /// where elements of the the top level are potentially nullable
        /// references to buckets of up-to 32 items in bottom level.
        /// </para>
        /// <para>
        /// Both level can increase size through doubling, through
        /// only the top one can be indefinite size.
        /// </para>
        /// </summary>
        private sealed class Lut<T>
        {
            private const int BottomLutBits = 5;
            private const int BottomLutMask = (1 << BottomLutBits) - 1;

            /// <summary>
            /// The default value lut ref returns for elements not defined in the lut.
            /// </summary>
            private static readonly T DefaultValue = default;

            /// <summary>
            /// A sparse array of values in the lut. The top level always allocated at least one element.
            /// </summary>
            private LutBucket[] _buckets = new LutBucket[1];

            /// <summary>
            /// Get maximal node that is used. Return -1 if LUT is unused.
            /// </summary>
            internal int MaxUsedIndex { get; private set; } = -1;

            /// <summary>
            /// Does LUT contains at least one used element?
            /// </summary>
            internal bool IsEmpty => MaxUsedIndex < 0;

            /// <summary>
            /// Get a value at specified index.
            /// </summary>
            /// <param name="index">Index, starting at 0.</param>
            /// <returns>Reference to an element at index, if the element is used, otherwise <see cref="DefaultValue"/>.</returns>
            internal ref readonly T Get(int index)
            {
                var (topIdx, bottomIdx) = SplitIndex(index);
                if (topIdx >= _buckets.Length)
                    return ref DefaultValue;

                if (!IsUsed(topIdx, bottomIdx))
                    return ref DefaultValue;

                var nodes = _buckets[topIdx].Nodes;
                return ref nodes[bottomIdx];
            }

            /// <summary>
            /// Does the index set a mask of used index (=was value set and not cleared)?
            /// </summary>
            internal bool IsUsed(int index)
            {
                var (topIdx, bottomIdx) = SplitIndex(index);
                if (topIdx >= _buckets.Length)
                    return false;

                return IsUsed(topIdx, bottomIdx);
            }

            /// <summary>
            /// Set/clar an element at index to a specified value.
            /// The used flag will be if the value is <c>default</c> or not.
            /// </summary>
            internal void Set(int index, T value)
            {
                var (topIdx, bottomIdx) = SplitIndex(index);

                SetValue(value, topIdx, bottomIdx);

                var valueIsDefault = EqualityComparer<T>.Default.Equals(value, DefaultValue);
                if (valueIsDefault)
                    ClearBitmap(topIdx, bottomIdx);
                else
                    SetBitmap(topIdx, bottomIdx);

                if (_buckets[topIdx].Bitmap == 0)
                    _buckets[topIdx] = new LutBucket(null, 0);

                RecalculateMaxIndex(index);
            }

            private void SetValue(T value, int topIdx, int bottomIdx)
            {
                var topSize = _buckets.Length;
                if (topIdx >= topSize)
                {
                    do
                    {
                        topSize *= 2;
                    } while (topIdx >= topSize);

                    Array.Resize(ref _buckets, topSize);
                }

                var bucket = _buckets[topIdx];
                var bottomBucketExists = bucket.Nodes is not null;
                if (!bottomBucketExists)
                {
                    var initialSize = 4;
                    while (bottomIdx >= initialSize)
                        initialSize *= 2;

                    _buckets[topIdx] = bucket = new LutBucket(new T[initialSize], 0);
                }
                else
                {
                    // Bottom exists, but might not be large enough
                    var bottomSize = bucket.Nodes.Length;
                    if (bottomIdx >= bottomSize)
                    {
                        do
                        {
                            bottomSize *= 2;
                        } while (bottomIdx >= bottomSize);

                        var bucketNodes = bucket.Nodes;
                        Array.Resize(ref bucketNodes, bottomSize);
                        _buckets[topIdx] = bucket = new LutBucket(bucketNodes, bucket.Bitmap);
                    }
                }

                bucket.Nodes[bottomIdx] = value;
            }

            private static (int TopLevelIndex, int BottomLevelIndex) SplitIndex(int index)
            {
                var topIdx = index >> BottomLutBits;
                var bottomIdx = index & BottomLutMask;
                return (topIdx, bottomIdx);
            }

            private bool IsUsed(int topIdx, int bottomIdx)
                => (_buckets[topIdx].Bitmap & (1 << bottomIdx)) != 0;

            private void SetBitmap(int topIdx, int bottomIdx)
                => _buckets[topIdx] = new LutBucket(_buckets[topIdx].Nodes, _buckets[topIdx].Bitmap | (uint)1 << bottomIdx);

            private void ClearBitmap(int topIdx, int bottomIdx)
                => _buckets[topIdx] = new LutBucket(_buckets[topIdx].Nodes, _buckets[topIdx].Bitmap & ~((uint)1 << bottomIdx));

            private void RecalculateMaxIndex(int index)
            {
                if (MaxUsedIndex <= index)
                    MaxUsedIndex = CalculateMaxIndex();
            }

            private int CalculateMaxIndex()
            {
                for (var bucketIdx = _buckets.Length - 1; bucketIdx >= 0; --bucketIdx)
                {
                    var bitmap = _buckets[bucketIdx].Bitmap;
                    if (bitmap != 0)
                    {
                        return (bucketIdx << BottomLutBits) + bitmap.GetHighestSetBit();
                    }
                }

                return -1;
            }

            /// <summary>
            /// A bucket of bottom layer of LUT. Each bucket has up-to 32 elements.
            /// </summary>
            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            private readonly struct LutBucket
            {
                public readonly T[] Nodes;

                /// <summary>
                /// <para>
                /// A bitmap array that indicates which nodes have a set/no-default values values
                /// (1 = value has been set and there is an element in the <see cref="_buckets"/>,
                /// 0 = value hasn't been set and <see cref="_buckets"/> might exist or not).
                /// If the element at some index is not is not set and lut is asked for a value,
                /// it should return <see cref="DefaultValue"/>.
                /// </para>
                /// <para>
                /// The length of the bitmap array is same as the <see cref="_buckets"/>, for each
                /// bottom level bucket, the element of index 0 in the bucket is represented by
                /// lowest bit, element 31 is represented by highest bit.
                /// </para>
                /// <para>
                /// This is useful to make a distinction between a node that is empty
                /// and a node that had it's value se to <see cref="DefaultValue"/>.
                /// </para>
                /// </summary>
                public readonly uint Bitmap;

                internal LutBucket(T[] nodes, uint bitmap)
                {
                    Nodes = nodes;
                    Bitmap = bitmap;
                }
            }

            /// <summary>
            /// Enumerator of LUT used values from low index to high.
            /// </summary>
            internal struct LutEnumerator
            {
                private readonly Lut<T> _lut;
                private readonly int _endIdx;
                private int _idx;

                /// <summary>
                /// Create a new enumerator from subset of elements.
                /// </summary>
                /// <param name="lut">Lookup table to traverse.</param>
                /// <param name="startIdx">First desired index, included.</param>
                /// <param name="endIdx">Last desired index, included.</param>
                internal LutEnumerator(Lut<T> lut, int startIdx, int endIdx)
                {
                    Debug.Assert(startIdx <= endIdx);
                    _lut = lut;
                    _idx = startIdx - 1;
                    _endIdx = endIdx;
                }

                public ref T Current => ref _lut._buckets[_idx >> BottomLutBits].Nodes[_idx & BottomLutMask];

                /// <summary>
                /// Index of current element in the LUT. Only valid, if enumerator is valid.
                /// </summary>
                public int Index => _idx;

                public bool MoveNext()
                {
                    var usedIndex = GetNextUsedIndexAtOrLater(_idx + 1);
                    if (usedIndex > _endIdx)
                        return false;

                    _idx = usedIndex;
                    return true;
                }

                private int GetNextUsedIndexAtOrLater(int index)
                {
                    var buckets = _lut._buckets;
                    var (topIdx, bottomIdx) = SplitIndex(index);

                    while (topIdx < buckets.Length)
                    {
                        var setBitIndex = buckets[topIdx].Bitmap.GetLowestSetBitAbove(bottomIdx);
                        if (setBitIndex >= 0)
                            return topIdx * 32 + setBitIndex;

                        ++topIdx;
                        bottomIdx = 0;
                    }

                    // We are the end of LUT
                    return int.MaxValue;
                }
            }

            /// <summary>
            /// Enumerator of LUT used values from high index to low index.
            /// </summary>
            internal struct ReverseLutEnumerator
            {
                private readonly Lut<T> _lut;
                private readonly int _startIdx;
                private int _idx;

                internal ReverseLutEnumerator(Lut<T> lut, int startIdx, int endIdx)
                {
                    Debug.Assert(startIdx <= endIdx);
                    _lut = lut;
                    _idx = endIdx + 1;
                    _startIdx = startIdx;
                }

                public ref T Current => ref _lut._buckets[_idx >> BottomLutBits].Nodes[_idx & BottomLutMask];

                public int Index => _idx;

                public bool MoveNext()
                {
                    var usedIndex = GetPrevIndexAtOrBefore(_idx - 1);
                    if (usedIndex < _startIdx)
                        return false;

                    _idx = usedIndex;
                    return true;
                }

                private int GetPrevIndexAtOrBefore(int index)
                {
                    var buckets = _lut._buckets;
                    var (topIdx, bottomIdx) = SplitIndex(index);
                    if (topIdx >= buckets.Length)
                    {
                        topIdx = buckets.Length - 1;
                        bottomIdx = 31;
                    }

                    while (topIdx >= 0)
                    {
                        var setBitIndex = buckets[topIdx].Bitmap.GetHighestSetBitBelow(bottomIdx);
                        if (setBitIndex >= 0)
                            return topIdx * 32 + setBitIndex;

                        --topIdx;
                        bottomIdx = 31;
                    }

                    return int.MinValue;
                }
            }
        }
    }
}
