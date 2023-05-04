#nullable disable

// Keep this file CodeMaid organised and cleaned

using System.Diagnostics;

namespace ClosedXML.Excel
{
    internal static class IntegerExtensions
    {
        public static bool Between(this int val, int from, int to)
        {
            return val >= from && val <= to;
        }

        /// <summary>
        /// Get index of highest set bit &lt;= to <paramref name="maximalIndex"/> or -1 if no such bit.
        /// </summary>
        internal static int GetHighestSetBitBelow(this uint value, int maximalIndex)
        {
            Debug.Assert(maximalIndex >= 0 && maximalIndex < 32);
            const uint highestBit = 0x80000000;
            value <<= 31 - maximalIndex;
            while (value != 0)
            {
                if ((value & highestBit) != 0)
                    return maximalIndex;
                value <<= 1;
                maximalIndex--;
            }

            return -1;
        }

        /// <summary>
        /// Get index of lowest set bit &gt;= to <paramref name="minimalIndex"/> or -1 if no such bit.
        /// </summary>
        internal static int GetLowestSetBitAbove(this uint value, int minimalIndex)
        {
            value >>= minimalIndex;
            while (value != 0)
            {
                if ((value & 1) == 1)
                    return minimalIndex;
                value >>= 1;
                minimalIndex++;
            }

            return -1;
        }

        /// <summary>
        /// Get highest set bit index or -1 if no bit is set.
        /// </summary>
        internal static int GetHighestSetBit(this uint value)
        {
            var highestSetBitIndex = -1;
            while (value != 0)
            {
                value >>= 1;
                highestSetBitIndex++;
            }

            return highestSetBitIndex;
        }
    }
}
