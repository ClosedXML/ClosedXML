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

        internal void SetValues(String address, String tooltip)
        {
            Tooltip = tooltip;
            if (address[0] == '.')
            {
                _externalAddress = new Uri(address, UriKind.Relative);
                IsExternal = true;
            }
            else
            {
                Uri uri;
                if(Uri.TryCreate(address, UriKind.Absolute, out uri))
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

        internal void SetValues(Uri uri, String tooltip)
        {
            Tooltip = tooltip;
            _externalAddress = uri;
            IsExternal = true;
        }

        internal void SetValues(IXLCell cell, String tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = cell.Address.ToString();
            IsExternal = false;
        }

        internal void SetValues(IXLRangeBase range, String tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = range.RangeAddress.ToString();
            IsExternal = false;
        }

        internal XLWorksheet Worksheet { get; set; }
    }
}
