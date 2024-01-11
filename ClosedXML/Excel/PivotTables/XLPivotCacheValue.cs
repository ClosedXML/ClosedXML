using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Represents a single value in a pivot cache record.
    /// </summary>
    internal readonly struct XLPivotCacheValue
    {
        /// <summary>
        /// A memory used to hold value of a <see cref="Type"/>. Its
        /// interpretation depends on the type. It doesn't hold value
        /// for strings directly, because GC doesn't allow aliasing
        /// same 8 bytes for number or references. For strings, it contains
        /// an index to a string storage array that is stored separately.
        /// </summary>
        private readonly double _value;

        private XLPivotCacheValue(XLPivotCacheValueType type, double value)
        {
            Type = type;
            _value = value;
        }

        internal XLPivotCacheValueType Type { get; }

        internal static XLPivotCacheValue ForMissing()
        {
            return new XLPivotCacheValue(XLPivotCacheValueType.Missing, 0);
        }

        internal static XLPivotCacheValue ForNumber(double number)
        {
            if (double.IsNaN(number) || double.IsInfinity(number))
                throw new ArgumentOutOfRangeException();

            return new XLPivotCacheValue(XLPivotCacheValueType.Number, number);
        }

        internal static XLPivotCacheValue ForBoolean(bool boolean)
        {
            return new XLPivotCacheValue(XLPivotCacheValueType.Boolean, boolean ? 1 : 0);
        }

        internal static XLPivotCacheValue ForError(XLError error)
        {
            return new XLPivotCacheValue(XLPivotCacheValueType.Error, (int)error);
        }

        internal static XLPivotCacheValue ForText(string text, List<string> storage)
        {
            var index = storage.Count;
            storage.Add(text);
            return new XLPivotCacheValue(XLPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
        }

        internal static XLPivotCacheValue ForText(string text, Dictionary<string, int> stringMap, List<string> storage)
        {
            if (!stringMap.TryGetValue(text, out var index))
            {
                index = storage.Count;
                storage.Add(text);
                stringMap.Add(text, index);
                return new XLPivotCacheValue(XLPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
            }

            return new XLPivotCacheValue(XLPivotCacheValueType.String, BitConverter.Int64BitsToDouble(index));
        }

        internal static XLPivotCacheValue ForDateTime(DateTime dateTime)
        {
            return new XLPivotCacheValue(XLPivotCacheValueType.DateTime, BitConverter.Int64BitsToDouble(dateTime.Ticks));
        }

        internal static XLPivotCacheValue ForIndex(uint index)
        {
            return new XLPivotCacheValue(XLPivotCacheValueType.Index, BitConverter.Int64BitsToDouble(index));
        }

        internal XLCellValue GetCellValue(List<string> stringStorage, XLPivotCacheSharedItems sharedItems)
        {
            switch (Type)
            {
                case XLPivotCacheValueType.Missing:
                    return Blank.Value;

                case XLPivotCacheValueType.Number:
                    return _value;

                case XLPivotCacheValueType.Boolean:
                    return _value != 0;

                case XLPivotCacheValueType.Error:
                    return (XLError)_value;

                case XLPivotCacheValueType.String:
                    return GetText(stringStorage);

                case XLPivotCacheValueType.DateTime:
                    return GetDateTime();

                case XLPivotCacheValueType.Index:
                    var intIndex = unchecked((uint)BitConverter.DoubleToInt64Bits(_value));
                    return sharedItems[intIndex];

                default:
                    throw new NotSupportedException();
            }
        }

        internal double GetNumber() => _value;

        internal Boolean GetBoolean()
        {
            return _value != 0;
        }

        internal XLError GetError()
        {
            return (XLError)_value;
        }

        internal string GetText(IReadOnlyList<string> stringStorage)
        {
            var stringIndex = unchecked((int)BitConverter.DoubleToInt64Bits(_value));
            return stringStorage[stringIndex];
        }

        internal DateTime GetDateTime()
        {
            var ticks = BitConverter.DoubleToInt64Bits(_value);
            return new DateTime(ticks);
        }

        internal uint GetIndex()
        {
            return unchecked((uint)BitConverter.DoubleToInt64Bits(_value));
        }
    }
}
