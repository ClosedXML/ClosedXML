#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A comparer of a scalar logic. Each comparer with it's logic can be accessed through a static property.
    /// </summary>
    internal class ScalarValueComparer : IComparer<ScalarValue>
    {
        private readonly StringComparer _stringComparer;

        /// <summary>
        /// Compare scalar values according to logic of "Sort" data in Excel, though texts are compared case insensitive.
        /// </summary>
        /// <remarks>
        /// Order is
        /// <list type="number">
        ///   <item>Type Number, from low to high</item>
        ///   <item>Type Text, from low to high (non-culture specific, ordinal compare)</item>
        ///   <item>Type Logical, <c>FALSE</c>, then <c>TRUE</c>.</item>
        ///   <item>Type Error, all error values are treated as equal (at least they don't change order).</item>
        ///   <item>Type Blank, all values are treated as equal.</item>
        /// </list>
        /// </remarks>
        public static ScalarValueComparer SortIgnoreCase { get; } = new(StringComparer.OrdinalIgnoreCase);

        private ScalarValueComparer(StringComparer stringComparer)
        {
            _stringComparer = stringComparer;
        }

        public int Compare(ScalarValue x, ScalarValue y)
        {
            var xTypeOrder = GetTypeOrder(in x);
            var yTypeOrder = GetTypeOrder(in y);
            var typeCompare = xTypeOrder - yTypeOrder;
            if (typeCompare != 0)
                return typeCompare;

            // Both types are same
            if (x.IsLogical)
                return x.GetLogical().CompareTo(y.GetLogical());
            if (x.IsNumber)
                return x.GetNumber().CompareTo(y.GetNumber());
            if (x.IsText)
                return _stringComparer.Compare(x.GetText(), y.GetText());

            // Blank and errors are always treated as equal
            return 0;

            static int GetTypeOrder(in ScalarValue value)
            {
                if (value.IsNumber) return 0;
                if (value.IsText) return 1;
                if (value.IsLogical) return 2;
                if (value.IsError) return 3;
                return 4; /* Blank */
            }
        }
    }
}
