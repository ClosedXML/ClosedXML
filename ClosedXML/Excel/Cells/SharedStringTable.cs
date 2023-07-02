using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class that holds all texts in a workbook. Each text can be either a simple
    /// <c>string</c> or a <see cref="XLImmutableRichText"/>.
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
        private readonly Dictionary<object, int> _reverseDict = new();

        /// <summary>
        /// Number of texts the table holds reference to.
        /// </summary>
        internal int Count => _table.Count - _freeIds.Count;

        /// <summary>
        /// Get a string for specified id. Doesn't matter if it is a plain text or a rich text. In both cases, return text.
        /// </summary>
        internal string this[int id]
        {
            get
            {
                var potentialText = _table[id].Text;
                if (potentialText is string text)
                    return text;

                if (potentialText is XLImmutableRichText richText)
                    return richText.Text;

                throw new ArgumentException($"Id {id} has no text.");
            }
        }

        /// <summary>
        /// The principle is that every entry is a text, but only some are rich text.
        /// This tries to get a rich text, if it is one. If it is just plain text, return null.
        /// </summary>
        internal XLImmutableRichText? GetRichText(int id)
        {
            var text = _table[id].Text;
            if (text is null)
                throw new ArgumentException($"Id {id} has no text.");

            return text as XLImmutableRichText;
        }

        /// <summary>
        /// Get id for a text and increase a number of references to the text by one.
        /// </summary>
        /// <returns>Id of a text in the SST.</returns>
        internal int IncreaseRef(string text) => IncreaseTextRef(text);

        /// <inheritdoc cref="IncreaseRef(string)"/>
        internal int IncreaseRef(XLImmutableRichText text) => IncreaseTextRef(text);

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

        private int IncreaseTextRef(object text)
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

        private int AddText(object text)
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
            /// <summary>
            /// Either a <c>string</c>, <c>XLImmutableRichText</c> or null if <c><see cref="RefCount"/> == 0</c>.
            /// </summary>
            internal readonly object? Text;

            /// <summary>
            /// How many objects (cells, pivot cache entries...) reference the text.
            /// </summary>
            internal readonly int RefCount;

            internal Entry(object? text, int refCount)
            {
                Text = text;
                RefCount = refCount;
            }
        }
    }
}
