#region

using System;
using System.Text;

#endregion

namespace ClosedXML.Excel
{
    internal class XLAlignment : IXLAlignment
    {
        private readonly IXLStylized _container;
        private XLAlignmentHorizontalValues _horizontal;
        private Int32 _indent;
        private Boolean _justifyLastLine;
        private XLAlignmentReadingOrderValues _readingOrder;
        private Int32 _relativeIndent;
        private Boolean _shrinkToFit;
        private Int32 _textRotation;
        private XLAlignmentVerticalValues _vertical;
        private Boolean _wrapText;

        public XLAlignment() : this(null, XLWorkbook.DefaultStyle.Alignment)
        {
        }

        public XLAlignment(IXLStylized container, IXLAlignment d = null)
        {
            _container = container;
            if (d == null) return;

            _horizontal = d.Horizontal;
            _vertical = d.Vertical;
            _indent = d.Indent;
            _justifyLastLine = d.JustifyLastLine;
            _readingOrder = d.ReadingOrder;
            _relativeIndent = d.RelativeIndent;
            _shrinkToFit = d.ShrinkToFit;
            _textRotation = d.TextRotation;
            _wrapText = d.WrapText;
        }

        #region IXLAlignment Members

        public XLAlignmentHorizontalValues Horizontal
        {
            get { return _horizontal; }
            set
            {
                SetStyleChanged();
                Boolean updateIndent = !(
                                            value == XLAlignmentHorizontalValues.Left
                                            || value == XLAlignmentHorizontalValues.Right
                                            || value == XLAlignmentHorizontalValues.Distributed
                                        );

                if (_container != null && !_container.UpdatingStyle)
                {
                    _container.Styles.ForEach(s =>
                                                  {
                                                      s.Alignment.Horizontal = value;
                                                      if (updateIndent) s.Alignment.Indent = 0;
                                                  });
                }
                else
                {
                    _horizontal = value;
                    if (updateIndent)
                        _indent = 0;
                }
            }
        }

        public XLAlignmentVerticalValues Vertical
        {
            get { return _vertical; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.Vertical = value);
                else
                    _vertical = value;
            }
        }

        public Int32 Indent
        {
            get { return _indent; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.Indent = value);
                else
                {
                    if (_indent != value)
                    {
                        if (_horizontal == XLAlignmentHorizontalValues.General)
                            _horizontal = XLAlignmentHorizontalValues.Left;

                        if (value > 0 && !(
                                              _horizontal == XLAlignmentHorizontalValues.Left
                                              || _horizontal == XLAlignmentHorizontalValues.Right
                                              || _horizontal == XLAlignmentHorizontalValues.Distributed
                                          ))
                        {
                            throw new ArgumentException(
                                "For indents, only left, right, and distributed horizontal alignments are supported.");
                        }

                        _indent = value;
                    }
                }
            }
        }

        public Boolean JustifyLastLine
        {
            get { return _justifyLastLine; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.JustifyLastLine = value);
                else
                    _justifyLastLine = value;
            }
        }

        public XLAlignmentReadingOrderValues ReadingOrder
        {
            get { return _readingOrder; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.ReadingOrder = value);
                else
                    _readingOrder = value;
            }
        }

        public Int32 RelativeIndent
        {
            get { return _relativeIndent; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.RelativeIndent = value);
                else
                    _relativeIndent = value;
            }
        }

        public Boolean ShrinkToFit
        {
            get { return _shrinkToFit; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.ShrinkToFit = value);
                else
                    _shrinkToFit = value;
            }
        }

        public Int32 TextRotation
        {
            get { return _textRotation; }
            set
            {
                SetStyleChanged();
                Int32 rotation = value;

                if (rotation != 255 && (rotation < -90 || rotation > 180))
                    throw new ArgumentException("TextRotation must be between -90 and 180 degrees, or 255.");

                if (rotation < 0)
                    rotation = 90 + (rotation * -1);

                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.TextRotation = rotation);
                else
                    _textRotation = rotation;
            }
        }

        public Boolean WrapText
        {
            get { return _wrapText; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.WrapText = value);
                else
                    _wrapText = value;
            }
        }

        public Boolean TopToBottom
        {
            get { return _textRotation == 255; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Alignment.TextRotation = value ? 255 : 0);
                else
                    _textRotation = value ? 255 : 0;
            }
        }


        public bool Equals(IXLAlignment other)
        {
            if (other == null)
                return false;

            var otherA = other as XLAlignment;
            if (otherA == null)
                return false;

            return
                _horizontal == otherA._horizontal
                && _vertical == otherA._vertical
                && _indent == otherA._indent
                && _justifyLastLine == otherA._justifyLastLine
                && _readingOrder == otherA._readingOrder
                && _relativeIndent == otherA._relativeIndent
                && _shrinkToFit == otherA._shrinkToFit
                && _textRotation == otherA._textRotation
                && _wrapText == otherA._wrapText
                ;
        }

        public IXLStyle SetHorizontal(XLAlignmentHorizontalValues value)
        {
            Horizontal = value;
            return _container.Style;
        }

        public IXLStyle SetVertical(XLAlignmentVerticalValues value)
        {
            Vertical = value;
            return _container.Style;
        }

        public IXLStyle SetIndent(Int32 value)
        {
            Indent = value;
            return _container.Style;
        }

        public IXLStyle SetJustifyLastLine()
        {
            JustifyLastLine = true;
            return _container.Style;
        }

        public IXLStyle SetJustifyLastLine(Boolean value)
        {
            JustifyLastLine = value;
            return _container.Style;
        }

        public IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value)
        {
            ReadingOrder = value;
            return _container.Style;
        }

        public IXLStyle SetRelativeIndent(Int32 value)
        {
            RelativeIndent = value;
            return _container.Style;
        }

        public IXLStyle SetShrinkToFit()
        {
            ShrinkToFit = true;
            return _container.Style;
        }

        public IXLStyle SetShrinkToFit(Boolean value)
        {
            ShrinkToFit = value;
            return _container.Style;
        }

        public IXLStyle SetTextRotation(Int32 value)
        {
            TextRotation = value;
            return _container.Style;
        }

        public IXLStyle SetWrapText()
        {
            WrapText = true;
            return _container.Style;
        }

        public IXLStyle SetWrapText(Boolean value)
        {
            WrapText = value;
            return _container.Style;
        }

        public IXLStyle SetTopToBottom()
        {
            TopToBottom = true;
            return _container.Style;
        }

        public IXLStyle SetTopToBottom(Boolean value)
        {
            TopToBottom = value;
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
            sb.Append(Horizontal);
            sb.Append("-");
            sb.Append(Vertical);
            sb.Append("-");
            sb.Append(Indent);
            sb.Append("-");
            sb.Append(JustifyLastLine);
            sb.Append("-");
            sb.Append(ReadingOrder);
            sb.Append("-");
            sb.Append(RelativeIndent);
            sb.Append("-");
            sb.Append(ShrinkToFit);
            sb.Append("-");
            sb.Append(TextRotation);
            sb.Append("-");
            sb.Append(WrapText);
            sb.Append("-");
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return Equals((XLAlignment)obj);
        }

        public override int GetHashCode()
        {
            return (Int32)Horizontal
                   ^ (Int32)Vertical
                   ^ Indent
                   ^ JustifyLastLine.GetHashCode()
                   ^ (Int32)ReadingOrder
                   ^ RelativeIndent
                   ^ ShrinkToFit.GetHashCode()
                   ^ TextRotation
                   ^ WrapText.GetHashCode();
        }
    }
}