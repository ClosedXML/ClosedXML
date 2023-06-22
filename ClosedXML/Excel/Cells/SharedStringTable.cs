using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class that holds all strings in a workbook.
    /// </summary>
    internal class SharedStringTable
    {
        /// <summary>
        /// Table of <c>Id</c> to text. Some ids are empty (<c>entry.RefCount = 0</c>) and
        /// are tracked in <see cref="_freeIds"/>.
        /// </summary>
        private readonly List<Entry> _table = new();

        /// <summary>
        /// List of indexes in <see cref="_table"/> that are unused.
        /// </summary>
        private readonly List<int> _freeIds = new();

        /// <summary>
        /// text -&gt; id
        /// </summary>
        private readonly Dictionary<string, int> _reverseDict = new(StringComparer.Ordinal);

        /// <summary>
        /// Number of texts the table holds reference to.
        /// </summary>
        internal int Count => _table.Count - _freeIds.Count;

        /// <summary>
        /// Get a string for specified id.
        /// </summary>
        internal string this[int id] => _table[id].Text ?? throw new ArgumentException($"Id {id} has no text.");

        /// <summary>
        /// Get id for a text and increase a number of references to the text by one.
        /// </summary>
        /// <returns>Id of a text in the SST.</returns>
        internal int IncreaseRef(string text)
        {
            if (!_reverseDict.TryGetValue(text, out var id))
            {
                id = AddText(text);
                _reverseDict.Add(text, id);
                return id;
            }

            var entry = _table[id];
            _table[id] = new Entry(entry.Text, entry.RefCount + 1);
            return id;
        }

        /// <summary>
        /// Decrease reference count of a text and free if necessary.
        /// </summary>
        internal void DecreaseRef(int id)
        {
            var entry = _table[id];
            if (entry.Text is null)
                throw new InvalidOperationException("Trying to release a text that doesn't have a reference.");

            if (entry.RefCount > 1)
            {
                _table[id] = new Entry(entry.Text, entry.RefCount - 1);
                return;
            }

            _table[id] = new Entry(null, 0);
            _freeIds.Add(id);
            _reverseDict.Remove(entry.Text);
        }

        private int AddText(string text)
        {
            if (_freeIds.Count > 0)
            {
                // List only changes size, not underlaying array, if last element is removed.
                var lastIndex = _freeIds.Count - 1;
                var id = _freeIds[lastIndex];
                _freeIds.RemoveAt(lastIndex);
                _table[id] = new Entry(text, 1);
                return id;
            }

            var lastTableIndex = _table.Count;
            _table.Add(new Entry(text, 1));
            return lastTableIndex;
        }

        private readonly struct Entry
        {
            internal readonly string? Text;
            internal readonly int RefCount;

            internal Entry(string? text, int refCount)
            {
                Text = text;
                RefCount = refCount;
            }
        }
    }
}
