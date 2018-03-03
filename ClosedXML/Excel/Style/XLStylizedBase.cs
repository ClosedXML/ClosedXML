using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Base class for any workbook element that has or may have a style.
    /// </summary>
    public abstract class XLStylizedBase : IXLStylized
    {
        #region Properties
        /// <summary>
        /// Read-only style property.
        /// </summary>
        public XLStyleValue StyleValue { get; private set; }

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
        protected virtual IEnumerable<XLStylizedBase> Children 
        {
            get
            {
                return RangesUsed.OfType<XLStylizedBase>();
            }
        }

        public abstract IXLRanges RangesUsed { get; }

        public abstract IEnumerable<IXLStyle> Styles { get; }
        #endregion Properties

        public XLStylizedBase(XLStyleValue styleValue)
        {
            StyleValue = styleValue;
        }

        #region Private methods
        private void SetStyle(IXLStyle value, bool propagate = false)
        {
            SetStyle(XLStyleValue.FromKey(XLStyle.GenerateKey(value)), propagate);
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
        #endregion Private methods
    }
}
