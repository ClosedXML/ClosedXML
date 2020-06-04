// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal static class TypeExtensions
    {
        public static Type GetUnderlyingType(this Type type)
        {
            return Nullable.GetUnderlyingType(type) ?? type;
        }

        public static bool IsNullableType(this Type type)
        {
            return Nullable.GetUnderlyingType(type) != null;
        }

        public static bool IsNumber(this Type type)
        {
            return type == typeof(sbyte)
                   || type == typeof(byte)
                   || type == typeof(short)
                   || type == typeof(ushort)
                   || type == typeof(int)
                   || type == typeof(uint)
                   || type == typeof(long)
                   || type == typeof(ulong)
                   || type == typeof(float)
                   || type == typeof(double)
                   || type == typeof(decimal);
        }

        public static bool IsSimpleType(this Type type)
        {
            return type.IsPrimitive
                   || type == typeof(String)
                   || type == typeof(DateTime)
                   || type == typeof(TimeSpan)
                   || type.IsNumber();
        }
    }
}
