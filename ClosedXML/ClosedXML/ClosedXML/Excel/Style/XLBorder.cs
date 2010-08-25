using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace ClosedXML.Excel.Style
{
    public class XLBorder: IXLBorder
    {
        IXLStylized container;
        public XLBorder(IXLStylized container, IXLBorder defaultBorder = null)
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

        private Color leftBorderColor;
        public Color LeftBorderColor
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

        private Color rightBorderColor;
        public Color RightBorderColor
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

        private Color topBorderColor;
        public Color TopBorderColor
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

        private Color bottomBorderColor;
        public Color BottomBorderColor
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

        private Color diagonalBorderColor;
        public Color DiagonalBorderColor
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
            return
                LeftBorder.ToString() + "-" +
                LeftBorderColor.ToString() + "-" +
                RightBorder.ToString() + "-" +
                RightBorderColor.ToString() + "-" +
                TopBorder.ToString() + "-" +
                TopBorderColor.ToString() + "-" +
                BottomBorder.ToString() + "-" +
                BottomBorderColor.ToString() + "-" +
                DiagonalBorder.ToString() + "-" +
                DiagonalBorderColor.ToString() + "-" +
                DiagonalUp.ToString() + "-" +
                DiagonalDown.ToString();

        }
    }
}
