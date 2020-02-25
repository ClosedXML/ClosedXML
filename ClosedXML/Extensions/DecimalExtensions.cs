// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal static class DecimalExtensions
    {
        public static Decimal SaveRound(this Decimal value)
        {
            return Math.Round(value, 6);
        }
    }
}
