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

        private IXLColor _patternBackgroundColor;
        private IXLColor _patternColor;
        private XLFillPatternValues _patternType;

        public IXLColor BackgroundColor
        {
            get { return _patternColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Fill.BackgroundColor = value);
                else
                {
                    _patternType = XLFillPatternValues.Solid;
                    _patternColor = new XLColor(value);
                    _patternBackgroundColor = new XLColor(value);

                    PatternTypeModified = true;
                    PatternColorModified = true;
                    PatternBackgroundColorModified = true;
                }
            }
        }

        public Boolean PatternColorModified;
        public IXLColor PatternColor
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
        public IXLColor PatternBackgroundColor
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

        public IXLStyle SetBackgroundColor(IXLColor value)
        {
            BackgroundColor = value;
            return _container.Style;
        }

        public IXLStyle SetPatternColor(IXLColor value)
        {
            PatternColor = value;
            return _container.Style;
        }

        public IXLStyle SetPatternBackgroundColor(IXLColor value)
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

        public XLFill(IXLStylized container, IXLFill defaultFill = null)
        {
            _container = container;
            if (defaultFill == null) return;
            _patternType = defaultFill.PatternType;
            _patternColor = new XLColor(defaultFill.PatternColor);
            _patternBackgroundColor = new XLColor(defaultFill.PatternBackgroundColor);
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