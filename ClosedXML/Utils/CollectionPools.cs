using System;
using System.Collections.Generic;

namespace ClosedXML.Utils;

internal static class CollectionPools
{
    internal static class StackPool<T>
    {
        public static readonly ObjectPool<Stack<T>> Shared = new(() => new Stack<T>(), s => s.Clear());
    }

    internal class ObjectPool<T> where T : class
    {
        private readonly Func<T> _create;
        private readonly Action<T> _reset;
        private readonly Stack<T> _items = new();

        public ObjectPool(Func<T> create, Action<T> reset)
        {
            _create = create;
            _reset = reset;
        }

        public T Rent() => _items.Count > 0 ? _items.Pop() : _create();

        public void Return(T item)
        {
            _reset(item);
            _items.Push(item);
        }
    }
}