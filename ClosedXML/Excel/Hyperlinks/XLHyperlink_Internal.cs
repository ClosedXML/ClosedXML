using System;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        internal XLHyperlink()
        {

        }

        internal XLHyperlink(XLHyperlink hyperlink)
        {
            _externalAddress = hyperlink._externalAddress;
            _internalAddress = hyperlink._internalAddress;
            Tooltip = hyperlink.Tooltip;
            IsExternal = hyperlink.IsExternal;
        }

        internal void SetValues(string address, string tooltip)
        {
            Tooltip = tooltip;
            if (address[0] == '.')
            {
                _externalAddress = new Uri(address, UriKind.Relative);
                IsExternal = true;
            }
            else
            {
                if (Uri.TryCreate(address, UriKind.Absolute, out var uri))
                {
                    _externalAddress = uri;
                    IsExternal = true;
                }
                else
                {
                    _internalAddress = address;
                    IsExternal = false;
                }
            }
        }

        internal void SetValues(Uri uri, string tooltip)
        {
            Tooltip = tooltip;
            _externalAddress = uri;
            IsExternal = true;
        }

        internal void SetValues(IXLCell cell, string tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = cell.Address.ToString(XLReferenceStyle.A1, true);
            IsExternal = false;
        }

        internal void SetValues(IXLRangeBase range, string tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = range.RangeAddress.ToString(XLReferenceStyle.A1, true);
            IsExternal = false;
        }

        internal XLWorksheet Worksheet { get; set; }
    }
}
