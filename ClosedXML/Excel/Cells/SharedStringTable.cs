using System;
using System.Collections.Generic;
using System.Diagnostics;

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
        private readonly Dictionary<Text, int> _reverseDict = new();

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
                var potentialText = _table[id].Text.Value;
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
            var text = _table[id].Text.Value;
            if (text is null)
                throw new ArgumentException($"Id {id} has no text.");

            return text as XLImmutableRichText;
        }

        /// <summary>
        /// Get id for a text and increase a number of references to the text by one.
        /// </summary>
        /// <returns>Id of a text in the SST.</returns>
        internal int IncreaseRef(string text, bool inline) => IncreaseTextRef(new Text(text, inline));

        /// <inheritdoc cref="IncreaseRef(string, bool)"/>
        internal int IncreaseRef(XLImmutableRichText text, bool inline) => IncreaseTextRef(new Text(text, inline));

        /// <summary>
        /// Decrease reference count of a text and free if necessary.
        /// </summary>
        internal void DecreaseRef(int id)
        {
            var entry = _table[id];
            if (entry.Text.Value is null)
                throw new InvalidOperationException("Trying to release a text that doesn't have a reference.");

            if (entry.RefCount > 1)
            {
                _table[id] = new Entry(entry.Text, entry.RefCount - 1);
                return;
            }

            _table[id] = new Entry(Text.Empty, 0);
            _freeIds.Add(id);
            _reverseDict.Remove(entry.Text);
        }

        /// <summary>
        /// Get a map that takes the actual string id and returns an continuous sequence (i.e. no gaps).
        /// If an id if free (no ref count), the id is mapped to -1.
        /// </summary>
        internal List<int> GetConsecutiveMap()
        {
            var map = new List<int>(_table.Count);
            var mappedStringId = 0;
            for (var i = 0; i < _table.Count; ++i)
            {
                var entry = _table[i];
                var isShared =
                    entry.RefCount > 0 && // Only used entry can be written to sst
                    !entry.Text.Inline;  // Inline texts shouldn't be written to sst
                map.Add(isShared ? mappedStringId++ : -1);
            }

            return map;
        }

        private int IncreaseTextRef(Text text)
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

        private int AddText(Text text)
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

        /// <summary>
        /// A struct to hold a text. It also needs a flag for inline/shared, because they have to be different
        /// in the table. If there was no inline/shared flag, there would be no way to easily determine whether
        /// a text should be written to sst or it should be inlined.
        /// </summary>
        [DebuggerDisplay("{Value} (Shared:{!Inline})")]
        private readonly struct Text : IEquatable<Text>
        {
            internal static readonly Text Empty = new(null, false);

            /// <summary>
            /// Either a <c>string</c>, <c>XLImmutableRichText</c> or null if <c><see cref="Entry.RefCount"/> == 0</c>.
            /// </summary>
            internal readonly object? Value;

            /// <summary>
            /// Must be as flag for inline string, so the default value is false => ShareString is true by default 
            /// </summary>
            internal readonly bool Inline;

            internal Text(object? value, bool inline)
            {
                Value = value;
                Inline = inline;
            }

            public override bool Equals(object obj) => obj is Text other && Equals(other);

            public bool Equals(Text other) => Equals(Value, other.Value) && Inline == other.Inline;

            public override int GetHashCode()
            {
                unchecked
                {
                    return ((Value is not null ? Value.GetHashCode() : 0) * 397) ^ Inline.GetHashCode();
                }
            }
        }

        [DebuggerDisplay("{Text.Value}:{RefCount} (Shared:{!Text.Inline})")]
        private readonly struct Entry
        {
            internal readonly Text Text;

            /// <summary>
            /// How many objects (cells, pivot cache entries...) reference the text.
            /// </summary>
            internal readonly int RefCount;

            internal Entry(Text text, int refCount)
            {
                Text = text;
                RefCount = refCount;
            }
        }
    }
}
