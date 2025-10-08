#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
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

        public abstract IEnumerable<IXLRange> RangesUsed { get; }

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

        void IXLStylized.ModifyStyle(Func<XLStyleKey, XLStyleKey> modification)
        {
            var children = GetChildrenRecursively(this)
                .GroupBy(child => child.StyleValue, ReferenceEqualityComparer<XLStyleValue>.Instance);

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
    }
}
