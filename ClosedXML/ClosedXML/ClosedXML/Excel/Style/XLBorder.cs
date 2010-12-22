using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace ClosedXML.Excel
{
    internal class XLBorder : IXLBorder
    {
        IXLStylized container;

        public XLBorder() : this(null, XLWorkbook.DefaultStyle.Border) { }

        public XLBorder(IXLStylized container, IXLBorder defaultBorder)
        {
            this.container = container;
            if (defaultBorder != null)
            {
                leftBorder = defaultBorder.LeftBorder;
                leftBorderColor = defaultBorder.LeftBorderColor;
                rightBorder = defaultBorder.RightBorder;
                rightBorderColor = defaultBorder.RightBorderColor;
                topBorder = defaultBorder.TopBorder;
                topBorderColor = defaultBorder.TopBorderColor;
                bottomBorder = defaultBorder.BottomBorder;
                bottomBorderColor = defaultBorder.BottomBorderColor;
                diagonalBorder = defaultBorder.DiagonalBorder;
                diagonalBorderColor = defaultBorder.DiagonalBorderColor;
                diagonalUp = defaultBorder.DiagonalUp;
                diagonalDown = defaultBorder.DiagonalDown;
            }
        }

        private XLBorderStyleValues leftBorder;
        public XLBorderStyleValues LeftBorder
        {
            get
            {
                return leftBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.LeftBorder = value);
                else
                    leftBorder = value;
            }
        }

        private XLColor leftBorderColor;
        public XLColor LeftBorderColor
        {
            get
            {
                return leftBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.LeftBorderColor = value);
                else
                    leftBorderColor = value;
            }
        }

        private XLBorderStyleValues rightBorder;
        public XLBorderStyleValues RightBorder
        {
            get
            {
                return rightBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.RightBorder = value);
                else
                    rightBorder = value;
            }
        }

        private XLColor rightBorderColor;
        public XLColor RightBorderColor
        {
            get
            {
                return rightBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.RightBorderColor = value);
                else
                    rightBorderColor = value;
            }
        }

        private XLBorderStyleValues topBorder;
        public XLBorderStyleValues TopBorder
        {
            get
            {
                return topBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.TopBorder = value);
                else
                    topBorder = value;
            }
        }

        private XLColor topBorderColor;
        public XLColor TopBorderColor
        {
            get
            {
                return topBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.TopBorderColor = value);
                else
                    topBorderColor = value;
            }
        }

        private XLBorderStyleValues bottomBorder;
        public XLBorderStyleValues BottomBorder
        {
            get
            {
                return bottomBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.BottomBorder = value);
                else
                    bottomBorder = value;
            }
        }

        private XLColor bottomBorderColor;
        public XLColor BottomBorderColor
        {
            get
            {
                return bottomBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.BottomBorderColor = value);
                else
                    bottomBorderColor = value;
            }
        }

        private XLBorderStyleValues diagonalBorder;
        public XLBorderStyleValues DiagonalBorder
        {
            get
            {
                return diagonalBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalBorder = value);
                else
                    diagonalBorder = value;
            }
        }

        private XLColor diagonalBorderColor;
        public XLColor DiagonalBorderColor
        {
            get
            {
                return diagonalBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalBorderColor = value);
                else
                    diagonalBorderColor = value;
            }
        }

        private Boolean diagonalUp;
        public Boolean DiagonalUp
        {
            get
            {
                return diagonalUp;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalUp = value);
                else
                    diagonalUp = value;
            }
        }

        private Boolean diagonalDown;
        public Boolean DiagonalDown
        {
            get
            {
                return diagonalDown;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalDown = value);
                else
                    diagonalDown = value;
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(LeftBorder.ToString());
            sb.Append("-");
            sb.Append(LeftBorderColor.ToString());
            sb.Append("-");
            sb.Append(RightBorder.ToString());
            sb.Append("-");
            sb.Append(RightBorderColor.ToString());
            sb.Append("-");
            sb.Append(TopBorder.ToString());
            sb.Append("-");
            sb.Append(TopBorderColor.ToString());
            sb.Append("-");
            sb.Append(BottomBorder.ToString());
            sb.Append("-");
            sb.Append(BottomBorderColor.ToString());
            sb.Append("-");
            sb.Append(DiagonalBorder.ToString());
            sb.Append("-");
            sb.Append(DiagonalBorderColor.ToString());
            sb.Append("-");
            sb.Append(DiagonalUp.ToString());
            sb.Append("-");
            sb.Append(DiagonalDown.ToString());
            return sb.ToString();
        }

        public bool Equals(IXLBorder other)
        {
            return
                   this.LeftBorder.Equals(other.LeftBorder)
                && this.LeftBorderColor.Equals(other.LeftBorderColor)
                && this.RightBorder.Equals(other.RightBorder)
                && this.RightBorderColor.Equals(other.RightBorderColor)
                && this.TopBorder.Equals(other.TopBorder)
                && this.TopBorderColor.Equals(other.TopBorderColor)
                && this.BottomBorder.Equals(other.BottomBorder)
                && this.BottomBorderColor.Equals(other.BottomBorderColor)
                && this.DiagonalBorder.Equals(other.DiagonalBorder)
                && this.DiagonalBorderColor.Equals(other.DiagonalBorderColor)
                && this.DiagonalUp.Equals(other.DiagonalUp)
                && this.DiagonalDown.Equals(other.DiagonalDown)
                ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLBorder)obj);
        }

        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                int hash = 17;
                hash = hash * 23 + LeftBorder.GetHashCode();
                hash = hash * 23 + LeftBorderColor.GetHashCode();
                hash = hash * 23 + RightBorder.GetHashCode();
                hash = hash * 23 + RightBorderColor.GetHashCode();
                hash = hash * 23 + TopBorder.GetHashCode();
                hash = hash * 23 + TopBorderColor.GetHashCode();
                hash = hash * 23 + BottomBorder.GetHashCode();
                hash = hash * 23 + BottomBorderColor.GetHashCode();
                hash = hash * 23 + DiagonalBorder.GetHashCode();
                hash = hash * 23 + DiagonalBorderColor.GetHashCode();
                hash = hash * 23 + DiagonalUp.GetHashCode();
                hash = hash * 23 + DiagonalDown.GetHashCode();
                return hash;
            }
        }
    }
}
