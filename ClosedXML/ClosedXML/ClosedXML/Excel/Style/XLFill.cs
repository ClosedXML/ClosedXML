using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFill : IXLFill
    {
        #region Properties

        public IXLColor BackgroundColor
        {
            get
            {
                return patternColor;
            }
            set
            {                
                if (container != null && !container.UpdatingStyle)
                {
                    container.Styles.ForEach(s => s.Fill.BackgroundColor = value);
                }
                else
                {
                    patternType = XLFillPatternValues.Solid;
                    patternColor = new XLColor(value);
                    patternBackgroundColor = new XLColor(value);
                }
            }
        }

        private IXLColor patternColor;
        public IXLColor PatternColor
        {
            get
            {
                return patternColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Fill.PatternColor = value);
                else
                    patternColor = value;
            }
        }

        private IXLColor patternBackgroundColor;
        public IXLColor PatternBackgroundColor
        {
            get
            {
                return patternBackgroundColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Fill.PatternBackgroundColor = value);
                else
                    patternBackgroundColor = value;
            }
        }

        private XLFillPatternValues patternType;
        public XLFillPatternValues PatternType
        {
            get
            {
                return patternType;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Fill.PatternType = value);
                else
                    patternType = value;
            }
        }

        public IXLStyle SetBackgroundColor(IXLColor value) { BackgroundColor = value; return container.Style; }
        public IXLStyle SetPatternColor(IXLColor value) { PatternColor = value; return container.Style; }
        public IXLStyle SetPatternBackgroundColor(IXLColor value) { PatternBackgroundColor = value; return container.Style; }
        public IXLStyle SetPatternType(XLFillPatternValues value) { PatternType = value; return container.Style; }


        #endregion

        #region Constructors

        public XLFill(): this(null, XLWorkbook.DefaultStyle.Fill)
        {          
        }

        IXLStylized container;
        public XLFill(IXLStylized container, IXLFill defaultFill = null)
        {
            this.container = container;
            if (defaultFill != null)
            {
                patternType = defaultFill.PatternType;
                patternColor = new XLColor(defaultFill.PatternColor);
                patternBackgroundColor = new XLColor(defaultFill.PatternBackgroundColor);
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(BackgroundColor.ToString());
            sb.Append("-");
            sb.Append(PatternType.ToString());
            sb.Append("-");
            sb.Append(PatternColor.ToString());
            return sb.ToString();
        }

        #endregion

        public bool Equals(IXLFill other)
        {
            return
                patternType == other.PatternType
            && patternColor.Equals(other.PatternColor)
            ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLFill)obj);
        }

        public override int GetHashCode()
        {
            return BackgroundColor.GetHashCode()
                ^ (Int32)PatternType
                ^ PatternColor.GetHashCode();
        }
    }
}
