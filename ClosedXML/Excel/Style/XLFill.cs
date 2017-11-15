using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLFill : IXLFill
    {
        #region IXLFill Members

        public bool Equals(IXLFill other)
        {
            return
                _patternType == other.PatternType
                && _backgroundColor.Equals(other.BackgroundColor)
                && _patternColor.Equals(other.PatternColor)
                ;
        }

        #endregion IXLFill Members

        private void SetStyleChanged()
        {
            if (_container != null) _container.StyleChanged = true;
        }

        public override bool Equals(object obj)
        {
            return Equals((XLFill)obj);
        }

        public override int GetHashCode()
        {
            return BackgroundColor.GetHashCode()
                   ^ (Int32)PatternType
                   ^ PatternColor.GetHashCode();
        }

        #region Properties

        private XLColor _backgroundColor;
        private XLColor _patternColor;
        private XLFillPatternValues _patternType;

        public XLColor BackgroundColor
        {
            get { return _backgroundColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.BackgroundColor = value);
                else
                {
                    // 4 ways of determining an "empty" color
                    if (new XLFillPatternValues[] { XLFillPatternValues.None, XLFillPatternValues.Solid }.Contains(_patternType)
                        && (_backgroundColor == null
                        || !_backgroundColor.HasValue
                        || _backgroundColor == XLColor.NoColor
                        || _backgroundColor.ColorType == XLColorType.Indexed && _backgroundColor.Indexed == 64))
                    {
                        _patternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                    }
                    _backgroundColor = value;
                }
            }
        }

        public XLColor PatternColor
        {
            get { return _patternColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.PatternColor = value);
                else
                {
                    _patternColor = value;
                }
            }
        }

        public XLFillPatternValues PatternType
        {
            get { return _patternType; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.PatternType = value);
                else
                {
                    _patternType = value;
                }
            }
        }

        public IXLStyle SetBackgroundColor(XLColor value)
        {
            BackgroundColor = value;
            return _container.Style;
        }

        public IXLStyle SetPatternColor(XLColor value)
        {
            PatternColor = value;
            return _container.Style;
        }

        public IXLStyle SetPatternType(XLFillPatternValues value)
        {
            PatternType = value;
            return _container.Style;
        }

        #endregion Properties

        #region Constructors

        private readonly IXLStylized _container;

        public XLFill() : this(null, XLWorkbook.DefaultStyle.Fill)
        {
        }

        public XLFill(IXLStylized container, IXLFill defaultFill = null, Boolean useDefaultModify = true)
        {
            _container = container;
            if (defaultFill == null) return;
            _patternType = defaultFill.PatternType;
            _backgroundColor = defaultFill.BackgroundColor;
            _patternColor = defaultFill.PatternColor;

            if (useDefaultModify)
            {
                var d = defaultFill as XLFill;
            }
        }

        #endregion Constructors

        #region Overridden

        public override string ToString()
        {
            switch (PatternType)
            {
                case XLFillPatternValues.None:
                    return "None";

                case XLFillPatternValues.Solid:
                    return string.Concat("Solid ", BackgroundColor.ToString());

                default:
                    return string.Concat(PatternType.ToString(), " pattern: ", PatternColor.ToString(), " on ", BackgroundColor.ToString());
            }
        }

        #endregion Overridden
    }
}
