using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFill : IXLFill
    {
        #region IXLFill Members

        public bool Equals(IXLFill other)
        {
            return
                _patternType == other.PatternType
                && _patternColor.Equals(other.PatternColor)
                && _patternBackgroundColor.Equals(other.PatternBackgroundColor)
                ;
        }

        #endregion

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

        private XLColor _patternBackgroundColor;
        private XLColor _patternColor;
        private XLFillPatternValues _patternType;

        public XLColor BackgroundColor
        {
            get { return _patternColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.BackgroundColor = value);
                else
                {
                    _patternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                    _patternColor = value;
                    _patternBackgroundColor = value;

                    PatternTypeModified = true;
                    PatternColorModified = true;
                    PatternBackgroundColorModified = true;
                }
            }
        }

        public Boolean PatternColorModified;
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
                    PatternColorModified = true;
                }
            }
        }

        public Boolean PatternBackgroundColorModified;
        public XLColor PatternBackgroundColor
        {
            get { return _patternBackgroundColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.PatternBackgroundColor = value);
                else
                {
                    _patternBackgroundColor = value;
                    PatternBackgroundColorModified = true;
                }
            }
        }

        public Boolean PatternTypeModified;
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
                    PatternTypeModified = true;
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

        public IXLStyle SetPatternBackgroundColor(XLColor value)
        {
            PatternBackgroundColor = value;
            return _container.Style;
        }

        public IXLStyle SetPatternType(XLFillPatternValues value)
        {
            PatternType = value;
            return _container.Style;
        }

        #endregion

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
            _patternColor = defaultFill.PatternColor;
            _patternBackgroundColor = defaultFill.PatternBackgroundColor;

            if (useDefaultModify)
            {
                var d = defaultFill as XLFill;
                PatternBackgroundColorModified = d.PatternBackgroundColorModified;
                PatternColorModified = d.PatternColorModified;
                PatternTypeModified = d.PatternTypeModified;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(BackgroundColor);
            sb.Append("-");
            sb.Append(PatternType.ToString());
            sb.Append("-");
            sb.Append(PatternColor);
            return sb.ToString();
        }

        #endregion
    }
}