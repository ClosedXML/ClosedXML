using System;
using System.Collections.Generic;
using System.Linq;
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
            private set { _value = XLBorderValue.FromKey(ref value); }
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

        public XLBorder(IXLStylized container, XLStyle style, XLBorderKey key) : this(container, style, XLBorderValue.FromKey(ref key))
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
                        k.TopBorder =
                        k.BottomBorder =
                        k.LeftBorder =
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
                        k.TopBorderColor =
                        k.BottomBorderColor =
                        k.LeftBorderColor =
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
                        k.TopBorder =
                        k.BottomBorder =
                        k.LeftBorder =
                        k.RightBorder = value;
                        return k;
                    });
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        using (new RestoreOutsideBorder(r))
                        {
                            foreach (var cell in r.Cells())
                            {
                                (cell.Style.Border as XLBorder)
                                    .Modify(k =>
                                    {
                                        k.TopBorder =
                                        k.BottomBorder =
                                        k.LeftBorder =
                                        k.RightBorder = value;
                                        return k;
                                    });
                            }
                        }
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
                        k.TopBorderColor =
                        k.BottomBorderColor =
                        k.LeftBorderColor =
                        k.RightBorderColor = value.Key;
                        return k;
                    });
                }
                else
                {
                    foreach (IXLRange r in _container.RangesUsed)
                    {
                        using (new RestoreOutsideBorder(r))
                        {
                            foreach (var cell in r.Cells())
                            {
                                (cell.Style.Border as XLBorder)
                                    .Modify(k =>
                                    {
                                        k.TopBorderColor =
                                        k.BottomBorderColor =
                                        k.LeftBorderColor =
                                        k.RightBorderColor = value.Key;
                                        return k;
                                    });
                            }
                        }
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
            get
            {
                var colorKey = Key.LeftBorderColor;
                return XLColor.FromKey(ref colorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

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
            get
            {
                var colorKey = Key.RightBorderColor;
                return XLColor.FromKey(ref colorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

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
            get
            {
                var colorKey = Key.TopBorderColor;
                return XLColor.FromKey(ref colorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

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
            get
            {
                var colorKey = Key.BottomBorderColor;
                return XLColor.FromKey(ref colorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

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
            get
            {
                var colorKey = Key.DiagonalBorderColor;
                return XLColor.FromKey(ref colorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

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

        /// <summary>
        /// Helper class that remembers outside border state before editing (in constructor) and restore afterwards (on disposing).
        /// It presumes that size of the range does not change during the editing, else it will fail.
        /// </summary>
        private class RestoreOutsideBorder : IDisposable
        {
            private readonly IXLRange _range;
            private readonly Dictionary<int, XLBorderKey> _topBorders;
            private readonly Dictionary<int, XLBorderKey> _bottomBorders;
            private readonly Dictionary<int, XLBorderKey> _leftBorders;
            private readonly Dictionary<int, XLBorderKey> _rightBorders;

            public RestoreOutsideBorder(IXLRange range)
            {
                _range = range ?? throw new ArgumentNullException(nameof(range));

                _topBorders = range.FirstRow().Cells().ToDictionary(
                    c => c.Address.ColumnNumber - range.RangeAddress.FirstAddress.ColumnNumber + 1,
                    c => (c.Style as XLStyle).Key.Border);

                _bottomBorders = range.LastRow().Cells().ToDictionary(
                    c => c.Address.ColumnNumber - range.RangeAddress.FirstAddress.ColumnNumber + 1,
                    c => (c.Style as XLStyle).Key.Border);

                _leftBorders = range.FirstColumn().Cells().ToDictionary(
                    c => c.Address.RowNumber - range.RangeAddress.FirstAddress.RowNumber + 1,
                    c => (c.Style as XLStyle).Key.Border);

                _rightBorders = range.LastColumn().Cells().ToDictionary(
                    c => c.Address.RowNumber - range.RangeAddress.FirstAddress.RowNumber + 1,
                    c => (c.Style as XLStyle).Key.Border);
            }

            // Used by Janitor.Fody
            private void DisposeManaged()
            {
                _topBorders.ForEach(kp => (_range.FirstRow().Cell(kp.Key).Style
                    .Border as XLBorder).Modify(k =>
                    {
                        k.TopBorder = kp.Value.TopBorder;
                        k.TopBorderColor = kp.Value.TopBorderColor;
                        return k;
                    }));
                _bottomBorders.ForEach(kp => (_range.LastRow().Cell(kp.Key).Style
                    .Border as XLBorder).Modify(k =>
                    {
                        k.BottomBorder = kp.Value.BottomBorder;
                        k.BottomBorderColor = kp.Value.BottomBorderColor;
                        return k;
                    }));
                _leftBorders.ForEach(kp => (_range.FirstColumn().Cell(kp.Key).Style
                    .Border as XLBorder).Modify(k =>
                    {
                        k.LeftBorder = kp.Value.LeftBorder;
                        k.LeftBorderColor = kp.Value.LeftBorderColor;
                        return k;
                    }));
                _rightBorders.ForEach(kp => (_range.LastColumn().Cell(kp.Key).Style
                    .Border as XLBorder).Modify(k =>
                    {
                        k.RightBorder = kp.Value.RightBorder;
                        k.RightBorderColor = kp.Value.RightBorderColor;
                        return k;
                    }));
            }

            public void Dispose()
            {
                // Leave this empty so that Janitor.Fody can do its work
            }
        }
    }
}
