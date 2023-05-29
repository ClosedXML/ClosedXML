// This file contains extensions methods that are present in .NET Core, but not in .NET Standard 2.0
#if !NETSTANDARD2_1_OR_GREATER
namespace System.IO
{
    internal static class StreamCompatibilityExtensions
    {
        public static int Read(this Stream stream, Span<byte> span)
        {
            for (var i = 0; i < span.Length; ++i)
            {
                var b = stream.ReadByte();
                if (b == -1)
                    return i;
                span[i] = (byte)b;
            }

            return span.Length;
        }
    }
}

namespace System
{
    public static class StringCompatibilityExtensions
    {
        public static bool Contains(this string s, char c)
        {
            return s.IndexOf(c) >= 0;
        }
    }
}

#endif
