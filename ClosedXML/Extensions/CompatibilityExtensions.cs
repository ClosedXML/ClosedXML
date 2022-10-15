// This file contains extensions methods that are present in .NET Core, but not in .NET Standard 2.0
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
