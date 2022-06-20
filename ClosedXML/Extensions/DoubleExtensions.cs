// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal static class DoubleExtensions
    {
        public static double SaveRound(this double value)
        {
            return Math.Round(value, 6);
        }
    }
}
