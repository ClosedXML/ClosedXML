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
        internal XLStyleValue StyleValue { get; private protected set; }
        XLStyleValue IXLStylized.StyleValue
        {
            get { return StyleValue; }
        }

        /// <summary>
        /// Editable style of the workbook element. Modification of this property DOES affect styles of child objects as well - they will
        /// be changed accordingly. Accessing this property causes a new <see cref="XLStyle"/> instance generated so use this property
        /// with caution. If you need only _read_ the style consider using <see cref="StyleValue"/> property instead.
        /// </summary>
        public IXLStyle Style
        {
            get { return InnerStyle; }
            set { SetStyle(value, true); }
        }

        /// <summary>
        /// Editable style of the workbook element. Modification of this property DOES NOT affect styles of child objects.
        /// Accessing this property causes a new <see cref="XLStyle"/> instance generated so use this property with caution. If you need
        /// only _read_ the style consider using <see cref="StyleValue"/> property instead.
        /// </summary>
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

        protected XLStylizedBase(XLStyleValue styleValue = null)
        {
            StyleValue = styleValue ?? XLWorkbook.DefaultStyleValue;
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

        private IEnumerable<XLStylizedBase> GetChildrenRecursively(XLStylizedBase parent)
        {
            return new List<XLStylizedBase> { parent }
                   .Union(parent.Children.Where(child => child != parent).SelectMany(child => GetChildrenRecursively(child)));
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
