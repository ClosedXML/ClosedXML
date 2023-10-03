using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Statistics about a <see cref="XLPivotCacheValues">pivot cache field
    /// values</see>. These statistics are available, even if cache field
    /// doesn't have any record values.
    /// </summary>
    internal readonly struct XLPivotCacheValuesStats
    {
        internal XLPivotCacheValuesStats(
            bool containsBlank,
            bool containsNumber,
            bool containsInteger,
            double? minValue,
            double? maxValue,
            bool containsString,
            bool longText,
            bool containsDate,
            DateTime? minDate,
            DateTime? maxDate)
        {
            ContainsBlank = containsBlank;
            ContainsNumber = containsNumber;
            ContainsInteger = containsInteger;
            MinValue = minValue;
            MaxValue = maxValue;
            ContainsString = containsString;
            LongText = longText;
            ContainsDate = containsDate;
            MinDate = minDate;
            MaxDate = maxDate;
        }

        internal bool ContainsBlank { get; }

        internal bool ContainsNumber { get; }

        /// <summary>
        /// Are all numbers in the field integers? Doesn't
        /// have to fit into int32/64, just no fractions.
        /// </summary>
        internal bool ContainsInteger { get; }

        internal double? MinValue { get; }

        internal double? MaxValue { get; }

        /// <summary>
        /// Does field contain any string, boolean or error?
        /// </summary>
        internal bool ContainsString { get; }

        /// <summary>
        /// Is any text longer than 255 chars?
        /// </summary>
        internal bool LongText { get; }

        /// <summary>
        /// Is any value <c>DateTime</c> or <c>TimeSpan</c>? TimeSpan is
        /// converted to <em>1899-12-31TXX:XX:XX</em> date.
        /// </summary>
        internal bool ContainsDate { get; }

        internal DateTime? MinDate { get; }

        internal DateTime? MaxDate { get; }
    }
}
