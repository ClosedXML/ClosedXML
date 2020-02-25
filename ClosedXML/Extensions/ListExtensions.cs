// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal static class ListExtensions
    {
        public static void RemoveAll<T>(this IList<T> list, Func<T, bool> predicate)
        {
            var indices = list.Where(item => predicate(item)).Select((item, i) => i).OrderByDescending(i => i).ToList();
            foreach (var i in indices)
            {
                list.RemoveAt(i);
            }
        }
    }
}
