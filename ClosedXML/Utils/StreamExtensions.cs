using System;
using System.IO;

namespace ClosedXML.Utils
{
    internal static class StreamExtensions
    {
        public static int ReadS32LE(this Stream stream)
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            var b3 = stream.ReadU8();
            var b4 = stream.ReadU8();
            return b4 << 24 | b3 << 16 | b2 << 8 | b1;
        }

        public static short ReadS16BE(this Stream stream)
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            return (short)((b1 << 8) | b2);
        }

        public static short ReadS16LE(this Stream stream)
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            return (short)((b2 << 8) | b1);
        }

        public static int ReadS32BE(this Stream stream)
        {
            var b1 = stream.ReadU8();
            var b2 = stream.ReadU8();
            var b3 = stream.ReadU8();
            var b4 = stream.ReadU8();
            return b1 << 24 | b2 << 16 | b3 << 8 | b4;
        }

        public static ushort ReadU16BE(this Stream stream)
        {
            if (!TryReadU16BE(stream, out var number))
                throw EndOfStreamException();
            return number;
        }

        public static uint ReadU32BE(this Stream stream)
        {
            if (!TryReadU32BE(stream, out var number))
                throw EndOfStreamException();
            return number;
        }

        public static uint ReadU32LE(this Stream stream)
        {
            if (!TryReadU32LE(stream, out var number))
                throw EndOfStreamException();
            return number;
        }

        public static bool TryReadU32LE(this Stream stream, out uint number)
        {
            if (!TryReadLE(stream, 4, out var result))
            {
                number = 0;
                return false;
            }
            number = (uint)result;
            return true;
        }

        public static ushort ReadU16LE(this Stream stream)
        {
            if (!TryReadU16LE(stream, out var number))
                throw EndOfStreamException();
            return number;
        }

        public static bool TryReadU16LE(this Stream stream, out ushort number)
        {
            if (!TryReadLE(stream, 2, out var result))
            {
                number = 0;
                return false;
            }
            number = (ushort)result;
            return true;
        }

        public static byte ReadU8(this Stream stream)
        {
            var b = stream.ReadByte();
            if (b == -1)
                throw EndOfStreamException();
            return (byte)b;
        }

        public static bool TryReadU32BE(this Stream stream, out uint number)
        {
            if (!TryReadBE(stream, 4, out var readNumber))
            {
                number = 0;
                return false;
            }
            number = (uint)readNumber;
            return true;
        }

        public static bool TryReadU16BE(this Stream stream, out ushort number)
        {
            int readNumber;
            if (TryReadBE(stream, 2, out readNumber))
            {
                number = (ushort)readNumber;
                return true;
            }
            number = default;
            return false;
        }

        private static bool TryReadLE(Stream stream, int size, out int number)
        {
            number = 0;
            for (var i = 0; i < size; ++i)
            {
                var readByte = stream.ReadByte();
                if (readByte == -1)
                    return false;

                number |= readByte << i * 8;
            }

            return true;
        }

        private static bool TryReadBE(Stream stream, int size, out int number)
        {
            number = 0;
            for (var i = 1; i <= size; ++i)
            {
                var readByte = stream.ReadByte();
                if (readByte == -1)
                    return false;

                number |= readByte << (size - i) * 8;
            }
            return true;
        }

        private static ArgumentException EndOfStreamException() => new("Unexpected end of stream.");
    }
}
