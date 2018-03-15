using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLBorder : IXLBorder
    {
        #region Static members

        internal static XLBorderKey GenerateKey(IXLBorder defaultBorder)
        {
            XLBorderKey key;
            if (defaultBorder == null)
            {
                key = XLBorderValue.Default.Key;
            }
            else if (defaultBorder is XLBorder)
            {
                key = (defaultBorder as XLBorder).Key;
            }
            else
            {
                key = new XLBorderKey
                {
                    LeftBorder = defaultBorder.LeftBorder,
                    LeftBorderColor = defaultBorder.LeftBorderColor.Key,
                    RightBorder = defaultBorder.RightBorder,
                    RightBorderColor = defaultBorder.RightBorderColor.Key,
                    TopBorder = defaultBorder.TopBorder,
                    TopBorderColor = defaultBorder.TopBorderColor.Key,
                    BottomBorder = defaultBorder.BottomBorder,
                    BottomBorderColor = defaultBorder.BottomBorderColor.Key,
                    DiagonalBorder = defaultBorder.DiagonalBorder,
                    DiagonalBorderColor = defaultBorder.DiagonalBorderColor.Key,
                    DiagonalUp = defaultBorder.DiagonalUp,
                    DiagonalDown = defaultBorder.DiagonalDown,
                };
            }
            return key;
        }

        #endregion Static members

        private readonly XLStyle _style;

        private readonly IXLStylized _container;

        private XLBorderValue _value;

        internal XLBorderKey Key
        {
            get { return _value.Key; }
            private set { _value = XLBorderValue.FromKey(value); }
        }

        #region Constructors

        /// <summary>
        /// Create an instance of XLBorder initializing it with the specified value.
        /// </summary>
        /// <param name="container">Container the border is applied to.</param>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLBorder(IXLStylized container, XLStyle style, XLBorderValue value)
        {
            _container = container;
            _style = style ?? _container.Style as XLStyle ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLBorder(IXLStylized container, XLStyle style, XLBorderKey key) : this(container, style, XLBorderValue.FromKey(key))
        {
        }

        public XLBorder(IXLStylized container, XLStyle style = null, IXLBorder d = null) : this(container, style, GenerateKey(d))
        {
        }

        #endregion Constructors

        #region IXLBorder Members

        public XLBorderStyleValues OutsideBorder
        {
            set
            {
                if (_container == null) return;

                if (_container is XLWorksheet || _container is XLConditionalFormat)
                {
                    Modify(k =>
                    {
                        k.TopBorder = value;
                        k.BottomBorder = value;
                        k.LeftBorder = value;
                        k.RightBorder = value;
                        return k;
                    });
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

        public XLColor OutsideBorderColor
        {
            set
            {
                if (_container == null) return;

                if (_container is XLWorksheet || _container is XLConditionalFormat)
                {
                    Modify(k =>
                    {
                        k.TopBorderColor = value.Key;
                        k.BottomBorderColor = value.Key;
                        k.LeftBorderColor = value.Key;
                        k.RightBorderColor = value.Key;
                        return k;
                    });
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
                if (_container == null) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    Modify(k =>
                    {
                        k.TopBorder = value;
                        k.BottomBorder = value;
                        k.LeftBorder = value;
                        k.RightBorder = value;
                        return k;
                    });
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

        public XLColor InsideBorderColor
        {
            set
            {
                if (_container == null) return;

                var wsContainer = _container as XLWorksheet;
                if (wsContainer != null)
                {
                    Modify(k =>
                    {
                        k.TopBorderColor = value.Key;
                        k.BottomBorderColor = value.Key;
                        k.LeftBorderColor = value.Key;
                        k.RightBorderColor = value.Key;
                        return k;
                    });
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        Dictionary<Int32, XLColor> topBorders = new Dictionary<int, XLColor>();
                        r.FirstRow().Cells().ForEach(
                            c =>
                            topBorders.Add(
                                c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                c.Style.Border.TopBorderColor));

                        Dictionary<Int32, XLColor> bottomBorders = new Dictionary<int, XLColor>();
                        r.LastRow().Cells().ForEach(
                            c =>
                            bottomBorders.Add(
                                c.Address.ColumnNumber - r.RangeAddress.FirstAddress.ColumnNumber + 1,
                                c.Style.Border.BottomBorderColor));

                        Dictionary<Int32, XLColor> leftBorders = new Dictionary<int, XLColor>();
                        r.FirstColumn().Cells().ForEach(
                            c =>
                            leftBorders.Add(
                                c.Address.RowNumber - r.RangeAddress.FirstAddress.RowNumber + 1,
                                c.Style.Border.LeftBorderColor));

                        Dictionary<Int32, XLColor> rightBorders = new Dictionary<int, XLColor>();
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
            get { return Key.LeftBorder; }
            set { Modify(k => { k.LeftBorder = value; return k; }); }
        }

        public XLColor LeftBorderColor
        {
            get { return XLColor.FromKey(Key.LeftBorderColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Color cannot be null");

                Modify(k => { k.LeftBorderColor = value.Key; return k; });
            }
        }

        public XLBorderStyleValues RightBorder
        {
            get { return Key.RightBorder; }
            set { Modify(k => { k.RightBorder = value; return k; }); }
        }

        public XLColor RightBorderColor
        {
            get { return XLColor.FromKey(Key.RightBorderColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Color cannot be null");

                Modify(k => { k.RightBorderColor = value.Key; return k; });
            }
        }

        public XLBorderStyleValues TopBorder
        {
            get { return Key.TopBorder; }
            set { Modify(k => { k.TopBorder = value; return k; }); }
        }

        public XLColor TopBorderColor
        {
            get { return XLColor.FromKey(Key.TopBorderColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Color cannot be null");

                Modify(k => { k.TopBorderColor = value.Key; return k; });
            }
        }

        public XLBorderStyleValues BottomBorder
        {
            get { return Key.BottomBorder; }
            set { Modify(k => { k.BottomBorder = value; return k; }); }
        }

        public XLColor BottomBorderColor
        {
            get { return XLColor.FromKey(Key.BottomBorderColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Color cannot be null");

                Modify(k => { k.BottomBorderColor = value.Key; return k; });
            }
        }

        public XLBorderStyleValues DiagonalBorder
        {
            get { return Key.DiagonalBorder; }
            set { Modify(k => { k.DiagonalBorder = value; return k; }); }
        }

        public XLColor DiagonalBorderColor
        {
            get { return XLColor.FromKey(Key.DiagonalBorderColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("Color cannot be null");

                Modify(k => { k.DiagonalBorderColor = value.Key; return k; });
            }
        }

        public Boolean DiagonalUp
        {
            get { return Key.DiagonalUp; }
            set { Modify(k => { k.DiagonalUp = value; return k; }); }
        }

        public Boolean DiagonalDown
        {
            get { return Key.DiagonalDown; }
            set { Modify(k => { k.DiagonalDown = value; return k; }); }
        }

        public IXLStyle SetOutsideBorder(XLBorderStyleValues value)
        {
            OutsideBorder = value;
            return _container.Style;
        }

        public IXLStyle SetOutsideBorderColor(XLColor value)
        {
            OutsideBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetInsideBorder(XLBorderStyleValues value)
        {
            InsideBorder = value;
            return _container.Style;
        }

        public IXLStyle SetInsideBorderColor(XLColor value)
        {
            InsideBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetLeftBorder(XLBorderStyleValues value)
        {
            LeftBorder = value;
            return _container.Style;
        }

        public IXLStyle SetLeftBorderColor(XLColor value)
        {
            LeftBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetRightBorder(XLBorderStyleValues value)
        {
            RightBorder = value;
            return _container.Style;
        }

        public IXLStyle SetRightBorderColor(XLColor value)
        {
            RightBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetTopBorder(XLBorderStyleValues value)
        {
            TopBorder = value;
            return _container.Style;
        }

        public IXLStyle SetTopBorderColor(XLColor value)
        {
            TopBorderColor = value;
            return _container.Style;
        }

        public IXLStyle SetBottomBorder(XLBorderStyleValues value)
        {
            BottomBorder = value;
            return _container.Style;
        }

        public IXLStyle SetBottomBorderColor(XLColor value)
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

        public IXLStyle SetDiagonalBorderColor(XLColor value)
        {
            DiagonalBorderColor = value;
            return _container.Style;
        }

        #endregion IXLBorder Members

        private void Modify(Func<XLBorderKey, XLBorderKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var border = styleKey.Border;
                styleKey.Border = modification(border);
                return styleKey;
            });
        }

        #region Overridden

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
            return Equals(obj as XLBorder);
        }

        public bool Equals(IXLBorder other)
        {
            var otherB = other as XLBorder;
            if (otherB == null)
                return false;

            return Key == otherB.Key;
        }

        public override int GetHashCode()
        {
            var hashCode = 416600561;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
