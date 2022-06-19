using System;
using System.Linq;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        private Uri _externalAddress;
        private string _internalAddress;

        public XLHyperlink(string address)
        {
            SetValues(address, string.Empty);
        }

        public XLHyperlink(string address, string tooltip)
        {
            SetValues(address, tooltip);
        }

        public XLHyperlink(IXLCell cell)
        {
            SetValues(cell, string.Empty);
        }

        public XLHyperlink(IXLCell cell, string tooltip)
        {
            SetValues(cell, tooltip);
        }

        public XLHyperlink(IXLRangeBase range)
        {
            SetValues(range, string.Empty);
        }

        public XLHyperlink(IXLRangeBase range, string tooltip)
        {
            SetValues(range, tooltip);
        }

        public XLHyperlink(Uri uri)
        {
            SetValues(uri, string.Empty);
        }

        public XLHyperlink(Uri uri, string tooltip)
        {
            SetValues(uri, tooltip);
        }

        public bool IsExternal { get; set; }

        public Uri ExternalAddress
        {
            get
            {
                return IsExternal ? _externalAddress : null;
            }
            set
            {
                _externalAddress = value;
                IsExternal = true;
            }
        }

        public IXLCell Cell { get; internal set; }

        public string InternalAddress
        {
            get
            {
                if (IsExternal)
                {
                    return null;
                }

                if (_internalAddress.Contains('!'))
                {
                    return _internalAddress[0] != '\''
                               ? string.Concat(
                                    _internalAddress
                                        .Substring(0, _internalAddress.IndexOf('!'))
                                        .EscapeSheetName(),
                                    '!',
                                    _internalAddress.Substring(_internalAddress.IndexOf('!') + 1))
                               : _internalAddress;
                }
                return string.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    _internalAddress);
            }
            set
            {
                _internalAddress = value;
                IsExternal = false;
            }
        }

        public string Tooltip { get; set; }

        public void Delete()
        {
            if (Cell == null)
            {
                return;
            }

            Worksheet.Hyperlinks.Delete(Cell.Address);
            if (Cell.Style.Font.FontColor.Equals(XLColor.FromTheme(XLThemeColor.Hyperlink)))
            {
                Cell.Style.Font.FontColor = Worksheet.StyleValue.Font.FontColor;
            }

            Cell.Style.Font.Underline = Worksheet.StyleValue.Font.Underline;
        }
    }
}
