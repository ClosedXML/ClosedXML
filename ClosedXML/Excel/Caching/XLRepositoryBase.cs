using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections;

namespace ClosedXML.Excel.Caching
{
    internal abstract class XLRepositoryBase : IXLRepository
    {
        public abstract void Clear();
    }

    internal abstract class XLRepositoryBase<TKey, TValue> : XLRepositoryBase, IXLRepository<TKey, TValue>
        where TKey : struct, IEquatable<TKey>
        where TValue : class
    {
        const int CONCURRENCY_LEVEL = 4;
        const int INITIAL_CAPACITY = 1000;

        private readonly ConcurrentDictionary<TKey, WeakReference<TValue>> _storage;
        private readonly Func<TKey, TValue> _createNew;

        protected XLRepositoryBase(Func<TKey, TValue> createNew)
            : this(createNew, EqualityComparer<TKey>.Default)
        {
        }

        protected XLRepositoryBase(Func<TKey, TValue> createNew, IEqualityComparer<TKey> comparer)
        {
            _storage = new ConcurrentDictionary<TKey, WeakReference<TValue>>(CONCURRENCY_LEVEL, INITIAL_CAPACITY, comparer);
            _createNew = createNew;
        }

        /// <summary>
        /// Check if the specified key is presented in the repository.
        /// </summary>
        /// <param name="key">Key to look for.</param>
        /// <returns>True if entry exists and alive, false otherwise.</returns>
        public bool ContainsKey(ref TKey key)
        {
            if (_storage.TryGetValue(key, out WeakReference<TValue> cachedReference))
            {
                if (cachedReference.TryGetTarget(out _))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Put the entity into the repository under the specified key if no other entity with
        /// the same key is presented.
        /// </summary>
        /// <param name="key">Key to identify the entity.</param>
        /// <param name="value">Entity to store.</param>
        /// <returns>Entity that is stored in the repository under the specified key
        /// (it can be either the <paramref name="value"/> or another entity that has been added to
        /// the repository before.)</returns>
        public TValue? Store(ref TKey key, TValue value)
        {
            if (value is null)
                return null;

            do
            {
                if (_storage.TryGetValue(key, out WeakReference<TValue> cachedReference) &&
                    cachedReference.TryGetTarget(out var storedValue))
                {
                    return storedValue;
                }
            } while (!_storage.TryAdd(key, new WeakReference<TValue>(value)));

            return value;
        }

        public TValue GetOrCreate(ref TKey key)
        {
            while (true)
            {
                // Try get existing weak ref
                if (_storage.TryGetValue(key, out var weakRef))
                {
                    // Try get the target value
                    if (weakRef.TryGetTarget(out var existingValue))
                    {
                        return existingValue;
                    }

                    // WeakReference target was collected, try replace (race safe)
                    var newValue = _createNew(key);
                    var newWeakRef = new WeakReference<TValue>(newValue);

                    // Update only if still the same stale weakRef to avoid overwriting another writer
                    if (_storage.TryUpdate(key, newWeakRef, weakRef))
                        return newValue;

                    // If failed, loop again (someone else replaced → race)
                    continue;
                }

                // Add new value (no existing weak ref)
                var createdValue = _createNew(key);
                var addedWeakRef = new WeakReference<TValue>(createdValue);

                if (_storage.TryAdd(key, addedWeakRef))
                    return createdValue;

                // If add failed → loop again as someone else inserted
            }
        }

        public TValue? Replace(ref TKey oldKey, ref TKey newKey)
        {
            if (_storage.TryRemove(oldKey, out WeakReference<TValue> cachedReference) && cachedReference != null)
            {
                _storage.TryAdd(newKey, cachedReference);
                return GetOrCreate(ref newKey);
            }

            return null;
        }

        public void Remove(ref TKey key)
        {
            _storage.TryRemove(key, out WeakReference<TValue> _);
        }

        public override void Clear()
        {
            _storage.Clear();
        }

        /// <summary>
        /// List items in the repository filtering out "dead" entries.
        /// </summary>
        public IEnumerator<TValue> GetEnumerator()
        {
            foreach (var pair in _storage)
            {
                if (pair.Value.TryGetTarget(out var value) && value != null)
                {
                    yield return value;
                }
                else
                {
                    _storage.TryRemove(pair.Key, out _);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}