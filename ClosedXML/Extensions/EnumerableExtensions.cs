// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T item in source)
                action(item);
        }

        public static Type? GetItemType(this IEnumerable source)
        {
            return GetGenericArgument(source.GetType());

            Type? GetGenericArgument(Type collectionType)
            {
                var ienumerable = collectionType.GetInterfaces()
                    .SingleOrDefault(i => i.GetGenericArguments().Length == 1 &&
                                          i.Name == "IEnumerable`1");

                return ienumerable?.GetGenericArguments()?.FirstOrDefault();
            }
        }

        public static HashSet<T> ToHashSet<T>(this IEnumerable<T> source)
        {
            return new HashSet<T>(source);
        }

        /// <summary>
        /// Skip last element of a sequence.
        /// </summary>
        public static IEnumerable<T> SkipLast<T>(this IEnumerable<T> source)
        {
            using var enumerator = source.GetEnumerator();
            if (!enumerator.MoveNext())
                yield break;

            T prev = enumerator.Current;
            while (enumerator.MoveNext())
            {
                yield return prev;
                prev = enumerator.Current;
            }
        }

        public static Boolean HasDuplicates<T>(this IEnumerable<T> source)
        {
            HashSet<T> distinctItems = new HashSet<T>();
            foreach (var item in source)
            {
                if (!distinctItems.Add(item))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Select all <typeparamref name="TItem"/> that are not null.
        /// </summary>
        public static IEnumerable<TItem> WhereNotNull<T, TItem>(this IEnumerable<T> source, Func<T, TItem?> property)
            where TItem : struct
        {
            return source.Select(property).Where(x => x.HasValue).Select(x => x!.Value);
        }
    }
}
