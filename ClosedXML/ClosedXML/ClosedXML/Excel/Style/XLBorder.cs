using System;
using System.Text;

namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLBorder : IXLBorder
    {
        private readonly IXLStylized _container;
        private XLBorderStyleValues _bottomBorder;
        private IXLColor _bottomBorderColor;
        private XLBorderStyleValues _diagonalBorder;
        private IXLColor _diagonalBorderColor;
        private Boolean _diagonalDown;
        private Boolean _diagonalUp;
        private XLBorderStyleValues _leftBorder;
        private IXLColor _leftBorderColor;
        private XLBorderStyleValues _rightBorder;
        private IXLColor _rightBorderColor;
        private XLBorderStyleValues _topBorder;
        private IXLColor _topBorderColor;

        public XLBorder() : this(null, XLWorkbook.DefaultStyle.Border)
        {
        }

        public XLBorder(IXLStylized container, IXLBorder defaultBorder)
        {
            _container = container;
            if (defaultBorder == null) return;

            _leftBorder = defaultBorder.LeftBorder;
            _leftBorderColor = new XLColor(defaultBorder.LeftBorderColor);
            _rightBorder = defaultBorder.RightBorder;
            _rightBorderColor = new XLColor(defaultBorder.RightBorderColor);
            _topBorder = defaultBorder.TopBorder;
            _topBorderColor = new XLColor(defaultBorder.TopBorderColor);
            _bottomBorder = defaultBorder.BottomBorder;
            _bottomBorderColor = new XLColor(defaultBorder.BottomBorderColor);
            _diagonalBorder = defaultBorder.DiagonalBorder;
            _diagonalBorderColor = new XLColor(defaultBorder.DiagonalBorderColor);
            _diagonalUp = defaultBorder.DiagonalUp;
            _diagonalDown = defaultBorder.DiagonalDown;
        }

        #region IXLBorder Members

        public XLBorderStyleValues OutsideBorder
        {
            set
            {
                if (_container == null || _container.UpdatingStyle) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    wsContainer.CellsUsed().Style.Border.SetOutsideBorder(value);
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        r.FirstColumn().Style.Border.LeftBorder = value;
                        r.LastColumn().Style.Border.RightBorder = value;
                        r.FirstRow().Style.Border.TopBorder = value;
                        r.LastRow().Style.Border.BottomBorder = value;
                    }
                }
            }
        }


        public IXLColor OutsideBorderColor
        {
            set
            {
                if (_container == null || _container.UpdatingStyle) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    wsContainer.CellsUsed().Style.Border.SetOutsideBorderColor(value);
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        r.FirstColumn().Style.Border.LeftBorderColor = value;
                        r.LastColumn().Style.Border.RightBorderColor = value;
                        r.FirstRow().Style.Border.TopBorderColor = value;
                        r.LastRow().Style.Border.BottomBorderColor = value;
                    }
                }
            }
        }

        public XLBorderStyleValues InsideBorder
        {
            set
            {
                if (_container == null || _container.UpdatingStyle) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    wsContainer.CellsUsed().Style.Border.SetOutsideBorder(value);
                    wsContainer.UpdatingStyle = true;
                    wsContainer.Style.Border.SetTopBorder(value);
                    wsContainer.Style.Border.SetBottomBorder(value);
                    wsContainer.Style.Border.SetLeftBorder(value);
                    wsContainer.Style.Border.SetRightBorder(value);
                    wsContainer.UpdatingStyle = false;
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        Dictionary<Int32, XLBorderStyleValues> topBorders = new Dictionary<int, XLBorderStyleValues>();
                        r.FirstRow().Cells().ForEach(
                            c =>
                            topBorders.Add(c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                           c.Style.Border.TopBorder));

                        Dictionary<Int32, XLBorderStyleValues> bottomBorders =
                            new Dictionary<int, XLBorderStyleValues>();
                        r.LastRow().Cells().ForEach(
                            c =>
                            bottomBorders.Add(c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                              c.Style.Border.BottomBorder));

                        Dictionary<Int32, XLBorderStyleValues> leftBorders = new Dictionary<int, XLBorderStyleValues>();
                        r.FirstColumn().Cells().ForEach(
                            c =>
                            leftBorders.Add(c.Address.RowNumber - r.RangeAddress.FirstAddress.RowNumber + 1,
                                            c.Style.Border.LeftBorder));

                        Dictionary<Int32, XLBorderStyleValues> rightBorders = new Dictionary<int, XLBorderStyleValues>();
                        r.LastColumn().Cells().ForEach(
                            c =>
                            rightBorders.Add(c.Address.RowNumber - r.RangeAddress.FirstAddress.RowNumber + 1,
                                             c.Style.Border.RightBorder));

                        r.Cells().Style.Border.OutsideBorder = value;

                        topBorders.ForEach(kp => r.FirstRow().Cell(kp.Key).Style.Border.TopBorder = kp.Value);
                        bottomBorders.ForEach(kp => r.LastRow().Cell(kp.Key).Style.Border.BottomBorder = kp.Value);
                        leftBorders.ForEach(kp => r.FirstColumn().Cell(kp.Key).Style.Border.LeftBorder = kp.Value);
                        rightBorders.ForEach(kp => r.LastColumn().Cell(kp.Key).Style.Border.RightBorder = kp.Value);
                    }
                }
            }
        }

        public IXLColor InsideBorderColor
        {
            set
            {
                if (_container == null || _container.UpdatingStyle) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    wsContainer.CellsUsed().Style.Border.SetOutsideBorderColor(value);
                    wsContainer.UpdatingStyle = true;
                    wsContainer.Style.Border.SetTopBorderColor(value);
                    wsContainer.Style.Border.SetBottomBorderColor(value);
                    wsContainer.Style.Border.SetLeftBorderColor(value);
                    wsContainer.Style.Border.SetRightBorderColor(value);
                    wsContainer.UpdatingStyle = false;
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        Dictionary<Int32, IXLColor> topBorders = new Dictionary<int, IXLColor>();
                        r.FirstRow().Cells().ForEach(
                            c =>
                            topBorders.Add(
                                c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                c.Style.Border.TopBorderColor));

                        Dictionary<Int32, IXLColor> bottomBorders = new Dictionary<int, IXLColor>();
                        r.LastRow().Cells().ForEach(
                            c =>
                            bottomBorders.Add(
                                c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                c.Style.Border.BottomBorderColor));

                        Dictionary<Int32, IXLColor> leftBorders = new Dictionary<int, IXLColor>();
                        r.FirstColumn().Cells().ForEach(
                            c =>
                            leftBorders.Add(
                                c.Address.RowNumber - r.RangeAddress.FirstAddress.RowNumber + 1,
                                c.Style.Border.LeftBorderColor));

                        Dictionary<Int32, IXLColor> rightBorders = new Dictionary<int, IXLColor>();
                        r.LastColumn().Cells().ForEach(
                            c =>
                            rightBorders.Add(
                                c.Address.RowNumber - r.RangeAddress.FirstAddress.RowNumber + 1,
                                c.Style.Border.RightBorderColor));

                        r.Cells().Style.Border.OutsideBorderColor = value;

                        topBorders.ForEach(
                            kp => r.FirstRow().Cell(kp.Key).Style.Border.TopBorderColor = kp.Value);
                        bottomBorders.ForEach(
                            kp => r.LastRow().Cell(kp.Key).Style.Border.BottomBorderColor = kp.Value);
                        leftBorders.ForEach(
                            kp => r.FirstColumn().Cell(kp.Key).Style.Border.LeftBorderColor = kp.Value);
                        rightBorders.ForEach(
                            kp => r.LastColumn().Cell(kp.Key).Style.Border.RightBorderColor = kp.Value);
                    }
                }
            }
        }

        public XLBorderStyleValues LeftBorder
        {
            get { return _leftBorder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.LeftBorder = value);
                else
                    _leftBorder = value;
            }
        }

        public IXLColor LeftBorderColor
        {
            get { return _leftBorderColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.LeftBorderColor = value);
                else
                    _leftBorderColor = value;
            }
        }

        public XLBorderStyleValues RightBorder
        {
            get { return _rightBorder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.RightBorder = value);
                else
                    _rightBorder = value;
            }
        }

        public IXLColor RightBorderColor
        {
            get { return _rightBorderColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.RightBorderColor = value);
                else
                    _rightBorderColor = value;
            }
        }

        public XLBorderStyleValues TopBorder
        {
            get { return _topBorder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.TopBorder = value);
                else
                    _topBorder = value;
            }
        }

        public IXLColor TopBorderColor
        {
            get { return _topBorderColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.TopBorderColor = value);
                else
                    _topBorderColor = value;
            }
        }

        public XLBorderStyleValues BottomBorder
        {
            get { return _bottomBorder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.BottomBorder = value);
                else
                    _bottomBorder = value;
            }
        }

        public IXLColor BottomBorderColor
        {
            get { return _bottomBorderColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.BottomBorderColor = value);
                else
                    _bottomBorderColor = value;
            }
        }

        public XLBorderStyleValues DiagonalBorder
        {
            get { return _diagonalBorder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.DiagonalBorder = value);
                else
                    _diagonalBorder = value;
            }
        }

        public IXLColor DiagonalBorderColor
        {
            get { return _diagonalBorderColor; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.DiagonalBorderColor = value);
                else
                    _diagonalBorderColor = value;
            }
        }

        public Boolean DiagonalUp
        {
            get { return _diagonalUp; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.DiagonalUp = value);
                else
                    _diagonalUp = value;
            }
        }

        public Boolean DiagonalDown
        {
            get { return _diagonalDown; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Border.DiagonalDown = value);
                else
                    _diagonalDown = value;
            }
        }

        public bool Equals(IXLBorder other)
        {
            var otherB = other as XLBorder;
            return
                _leftBorder == otherB._leftBorder
                && _leftBorderColor.Equals(otherB._leftBorderColor)
                && _rightBorder == otherB._rightBorder
                && _rightBorderColor.Equals(otherB._rightBorderColor)
                && _topBorder == otherB._topBorder
                && _topBorderColor.Equals(otherB._topBorderColor)
                && _bottomBorder == otherB._bottomBorder
                && _bottomBorderColor.Equals(otherB._bottomBorderColor)
                && _diagonalBorder == otherB._diagonalBorder
                && _diagonalBorderColor.Equals(otherB._diagonalBorderColor)
                && _diagonalUp == otherB._diagonalUp
                && _diagonalDown == otherB._diagonalDown
                ;
        }

        public IXLStyle SetOutsideBorder(XLBorderStyleValues value)
        {
            OutsideBorder = value;
            return _container.Style;
        }

        public IXLStyle SetOutsideBorderColor(IXLColor value)
        {
            OutsideBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetInsideBorder(XLBorderStyleValues value)
        {
            InsideBorder = value;
            return _container.Style;
        }

        public IXLStyle SetInsideBorderColor(IXLColor value)
        {
            InsideBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetLeftBorder(XLBorderStyleValues value)
        {
            LeftBorder = value;
            return _container.Style;
        }

        public IXLStyle SetLeftBorderColor(IXLColor value)
        {
            LeftBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetRightBorder(XLBorderStyleValues value)
        {
            RightBorder = value;
            return _container.Style;
        }

        public IXLStyle SetRightBorderColor(IXLColor value)
        {
            RightBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetTopBorder(XLBorderStyleValues value)
        {
            TopBorder = value;
            return _container.Style;
        }

        public IXLStyle SetTopBorderColor(IXLColor value)
        {
            TopBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetBottomBorder(XLBorderStyleValues value)
        {
            BottomBorder = value;
            return _container.Style;
        }

        public IXLStyle SetBottomBorderColor(IXLColor value)
        {
            BottomBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetDiagonalUp()
        {
            DiagonalUp = true;
            return _container.Style;
        }

        public IXLStyle SetDiagonalUp(Boolean value)
        {
            DiagonalUp = value;
            return _container.Style;
        }

        public IXLStyle SetDiagonalDown()
        {
            DiagonalDown = true;
            return _container.Style;
        }

        public IXLStyle SetDiagonalDown(Boolean value)
        {
            DiagonalDown = value;
            return _container.Style;
        }

        public IXLStyle SetDiagonalBorder(XLBorderStyleValues value)
        {
            DiagonalBorder = value;
            return _container.Style;
        }

        public IXLStyle SetDiagonalBorderColor(IXLColor value)
        {
            DiagonalBorderColor = value;
            return _container.Style;
        }

        #endregion

        private void SetStyleChanged()
        {
            if (_container != null) _container.StyleChanged = true;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(LeftBorder.ToString());
            sb.Append("-");
            sb.Append(LeftBorderColor);
            sb.Append("-");
            sb.Append(RightBorder.ToString());
            sb.Append("-");
            sb.Append(RightBorderColor);
            sb.Append("-");
            sb.Append(TopBorder.ToString());
            sb.Append("-");
            sb.Append(TopBorderColor);
            sb.Append("-");
            sb.Append(BottomBorder.ToString());
            sb.Append("-");
            sb.Append(BottomBorderColor);
            sb.Append("-");
            sb.Append(DiagonalBorder.ToString());
            sb.Append("-");
            sb.Append(DiagonalBorderColor);
            sb.Append("-");
            sb.Append(DiagonalUp.ToString());
            sb.Append("-");
            sb.Append(DiagonalDown.ToString());
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return Equals((XLBorder)obj);
        }

        public override int GetHashCode()
        {
            return (Int32)LeftBorder
                   ^ LeftBorderColor.GetHashCode()
                   ^ (Int32)RightBorder
                   ^ RightBorderColor.GetHashCode()
                   ^ (Int32)TopBorder
                   ^ TopBorderColor.GetHashCode()
                   ^ (Int32)BottomBorder
                   ^ BottomBorderColor.GetHashCode()
                   ^ (Int32)DiagonalBorder
                   ^ DiagonalBorderColor.GetHashCode()
                   ^ DiagonalUp.GetHashCode()
                   ^ DiagonalDown.GetHashCode();
        }
    }
}