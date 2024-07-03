#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Base class for any workbook element that has or may have a style.
    /// </summary>
    internal abstract class XLStylizedBase : IXLStylized
    {
        #region Properties

        /// <summary>
        /// Read-only style property.
        /// </summary>
        internal virtual XLStyleValue StyleValue { get; private protected set; }

        /// <inheritdoc cref="IXLStylized.StyleValue"/>
        XLStyleValue IXLStylized.StyleValue
        {
            get { return StyleValue; }
        }

        /// <inheritdoc cref="IXLStylized.Style"/>
        public IXLStyle Style
        {
            get { return InnerStyle; }
            set { SetStyle(value, true); }
        }

        /// <inheritdoc cref="IXLStylized.InnerStyle"/>
        public IXLStyle InnerStyle
        {
            get { return new XLStyle(this, StyleValue.Key); }
            set { SetStyle(value, false); }
        }

        /// <summary>
        /// Get a collection of stylized entities which current entity's style changes should be propagated to.
        /// </summary>
        protected abstract IEnumerable<XLStylizedBase> Children { get; }

        public abstract IXLRanges RangesUsed { get; }

        public abstract IEnumerable<IXLStyle> Styles { get; }

        #endregion Properties

        protected XLStylizedBase(XLStyleValue styleValue)
        {
            StyleValue = styleValue ?? XLWorkbook.DefaultStyleValue;
        }

        protected XLStylizedBase()
        {
            // Ctor only for XLCell that stores `StyleValue` in a slice. 
        }

        #region Private methods

        private void SetStyle(IXLStyle style, bool propagate = false)
        {
            if (style is XLStyle xlStyle)
                SetStyle(xlStyle.Value, propagate);
            else
            {
                var styleKey = XLStyle.GenerateKey(style);
                SetStyle(XLStyleValue.FromKey(ref styleKey), propagate);
            }
        }

        /// <summary>
        /// Apply specified style to the container.
        /// </summary>
        /// <param name="value">Style to apply.</param>
        /// <param name="propagate">Whether or not propagate the style to inner ranges.</param>
        private void SetStyle(XLStyleValue value, bool propagate = false)
        {
            StyleValue = value;
            if (propagate)
            {
                Children.ForEach(child => child.SetStyle(StyleValue, true));
            }
        }

        private static ReferenceEqualityComparer<XLStyleValue> _comparer = new ReferenceEqualityComparer<XLStyleValue>();

        void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
        {
            var children = GetChildrenRecursively(this)
                .GroupBy(child => child.StyleValue, _comparer);

            foreach (var group in children)
            {
                var styleKey = modification(group.Key.Key);
                var styleValue = XLStyleValue.FromKey(ref styleKey);
                foreach (var child in group)
                {
                    child.StyleValue = styleValue;
                }
            }
        }

        private static HashSet<XLStylizedBase> GetChildrenRecursively(XLStylizedBase parent)
        {
            void Collect(XLStylizedBase root, HashSet<XLStylizedBase> collector)
            {
                collector.Add(root);
                foreach (var child in root.Children)
                {
                    Collect(child, collector);
                }
            }

            var results = new HashSet<XLStylizedBase>();
            Collect(parent, results);

            return results;
        }

        #endregion Private methods

        #region Nested classes

        public sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
        {
            public bool Equals(T x, T y) => ReferenceEquals(x, y);

            public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
        }

        #endregion Nested classes
    }
}
