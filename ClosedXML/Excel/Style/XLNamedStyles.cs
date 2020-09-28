using System;
using System.Collections.Concurrent;

namespace ClosedXML.Excel
{
    internal class XLNamedStyles
    {
        private readonly ConcurrentDictionary<string, XLStyleKey> _keyByName;

        public XLNamedStyles()
        {
            _keyByName = new ConcurrentDictionary<string, XLStyleKey>(StringComparer.InvariantCultureIgnoreCase);
        }

        public void Add(string styleName, XLStyleValue style)
        {
            Add(styleName, style.Key);
        }

        public void Add(string styleName, XLStyleKey styleKey)
        {
            _keyByName.TryAdd(styleName, styleKey);
        }

        public XLStyleValue this[string name]
        {
            get
            {
                if (_keyByName.TryGetValue(name, out var styleKey))
                {
                    return XLStyleValue.FromKey(styleKey);
                }

                return null;
            }
        }
    }
}
