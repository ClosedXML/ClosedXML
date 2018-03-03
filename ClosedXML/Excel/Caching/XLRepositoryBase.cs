using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Collections;

namespace ClosedXML.Excel.Caching
{
    internal abstract class XLRepositoryBase : IXLRepository
    {
        public abstract void Clear();
    }

    internal abstract class XLRepositoryBase<Tkey, Tvalue> : XLRepositoryBase, IXLRepository<Tkey, Tvalue>
        where Tkey : struct, IEquatable<Tkey>
        where Tvalue : class
    {
        const int CONCURRENCY_LEVEL = 4;
        const int INITIAL_CAPACITY = 1000;

        private readonly ConcurrentDictionary<Tkey, WeakReference> _storage;
        private readonly Func<Tkey, Tvalue> _createNew;
        public XLRepositoryBase(Func<Tkey, Tvalue> createNew) : this(createNew, EqualityComparer<Tkey>.Default)
        {
        }

        public XLRepositoryBase(Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer)
        {
            _storage = new ConcurrentDictionary<Tkey, WeakReference>(CONCURRENCY_LEVEL, INITIAL_CAPACITY, comparer);
            _createNew = createNew;
        }

        public bool ContainsKey(Tkey key)
        {
            WeakReference cachedReference;
            if (_storage.TryGetValue(key, out cachedReference))
            {
                var storedValue = cachedReference.Target as Tvalue;
                return (storedValue != null);
            }
            return false;
        }

        public Tvalue Store(Tkey key, Tvalue value)
        {
            if (value == null)
                return null;

            if (!_storage.ContainsKey(key))
            {
                _storage.TryAdd(key, new WeakReference(value));
                return value;
            }
            else
            {
                var cachedReference = _storage[key];
                var storedValue = cachedReference.Target as Tvalue;
                if (storedValue == null)
                {
                    _storage.TryAdd(key, new WeakReference(value));
                    return value;
                }
                return storedValue;
            }
        }

        public Tvalue GetOrCreate(Tkey key)
        {
            WeakReference cachedReference;
            if (_storage.TryGetValue(key, out cachedReference))
            {
                var storedValue = cachedReference.Target as Tvalue;
                if (storedValue != null)
                {
                    return storedValue;
                }
                else
                {
                    WeakReference _;
                    _storage.TryRemove(key, out _);
                }
            }
            
            var value = _createNew(key);
            return Store(key, value);
        }

        public override void Clear()
        {
            _storage.Clear();
        }

        /// <summary>
        /// Enumerate items in repository removing "dead" entries.
        /// </summary>
        public IEnumerator<Tvalue> GetEnumerator()
        {
            return _storage
                .Select(pair =>
                {
                    var val = pair.Value.Target as Tvalue;
                    if (val == null)
                    {
                        WeakReference _;
                        _storage.TryRemove(pair.Key, out _);
                    }
                    return val;
                })
                .Where(val => val != null)
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
