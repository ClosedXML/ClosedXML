using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML
{
    public static class Extensions
    {
        // Adds the .ForEach method to all IEnumerables
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T item in source)
                action(item);
        }
    }
}
