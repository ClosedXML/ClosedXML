// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal static class DoubleExtensions
    {
        public static Double SaveRound(this Double value)
        {
            return Math.Round(value, 6);
        }
    }
}
