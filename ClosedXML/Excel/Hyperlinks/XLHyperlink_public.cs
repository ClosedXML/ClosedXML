using ClosedXML.Extensions;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        private Uri _externalAddress;
        private String _internalAddress;

        public XLHyperlink(String address)
        {
            SetValues(address, String.Empty);
        }

        public XLHyperlink(String address, String tooltip)
        {
            SetValues(address, tooltip);
        }

        public XLHyperlink(IXLCell cell)
        {
            SetValues(cell, String.Empty);
        }

        public XLHyperlink(IXLCell cell, String tooltip)
        {
            SetValues(cell, tooltip);
        }

        public XLHyperlink(IXLRangeBase range)
        {
            SetValues(range, String.Empty);
        }

        public XLHyperlink(IXLRangeBase range, String tooltip)
        {
            SetValues(range, tooltip);
        }

        public XLHyperlink(Uri uri)
        {
            SetValues(uri, String.Empty);
        }

        public XLHyperlink(Uri uri, String tooltip)
        {
            SetValues(uri, tooltip);
        }

        public Boolean IsExternal { get; set; }

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

        public String InternalAddress
        {
            get
            {
                if (IsExternal)
                    return null;
                if (_internalAddress.Contains('!'))
                {
                    return _internalAddress[0] != '\''
                               ? String.Concat(
                                    _internalAddress
                                        .Substring(0, _internalAddress.IndexOf('!'))
                                        .EscapeSheetName(),
                                    '!',
                                    _internalAddress.Substring(_internalAddress.IndexOf('!') + 1))
                               : _internalAddress;
                }
                return String.Concat(
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

        public String Tooltip { get; set; }

        public void Delete()
        {
            if (Cell == null) return;
            Worksheet.Hyperlinks.Delete(Cell.Address);
            if (Cell.Style.Font.FontColor.Equals(XLColor.FromTheme(XLThemeColor.Hyperlink)))
                Cell.Style.Font.FontColor = Worksheet.StyleValue.Font.FontColor;

            Cell.Style.Font.Underline = Worksheet.StyleValue.Font.Underline;
        }
    }
}
