#nullable disable

using System;
using System.Buffers;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using ClosedXML.Utils;

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
        /// <param name="propagate">Whether to propagate the style to inner ranges.</param>
        private void SetStyle(XLStyleValue value, bool propagate = false)
        {
            StyleValue = value;
            if (propagate)
            {
                Children.ForEach(child => child.SetStyle(StyleValue, true));
            }
        }

        private static readonly ReferenceEqualityComparer<XLStyleValue> _comparer = new();

        void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
        {
            var children = CollectChildrenRecursively(this);

            var groups = new Dictionary<XLStyleValue, List<XLStylizedBase>>(_comparer);
            foreach (var child in children)
            {
                if (!groups.TryGetValue(child.StyleValue, out var list))
                    groups[child.StyleValue] = list = new List<XLStylizedBase>();

                list.Add(child);
            }

            // Apply style modification
            foreach (var kvp in groups)
            {
                var originalStyleValue = kvp.Key;
                var modifiedKey = modification(originalStyleValue.Key);
                var modifiedStyleValue = XLStyleValue.FromKey(ref modifiedKey);

                foreach (var child in kvp.Value)
                {
                    child.StyleValue = modifiedStyleValue;
                }
            }
        }

        private static XLStylizedBase[] CollectChildrenRecursively(XLStylizedBase parent)
        {
            var stack = CollectionPools.StackPool<XLStylizedBase>.Shared.Rent();
            stack.Push(parent);

            var result = ArrayPool<XLStylizedBase>.Shared.Rent(4096); // initial guess
            var count = 0;

            while (stack.Count > 0)
            {
                var current = stack.Pop();

                if (count == result.Length)
                    Array.Resize(ref result, result.Length * 2);

                result[count++] = current;

                foreach (var child in current.Children)
                    stack.Push(child);
            }

            CollectionPools.StackPool<XLStylizedBase>.Shared.Return(stack);

            // If count < result.Length → trim (optional if needed)
            if (count < result.Length)
            {
                var trimmed = new XLStylizedBase[count];
                Array.Copy(result, trimmed, count);
                ArrayPool<XLStylizedBase>.Shared.Return(result);
                return trimmed;
            }

            return result;
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