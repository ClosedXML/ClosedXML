using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLHyperlinks: IXLHyperlinks
    {
        private readonly Dictionary<IXLAddress, XLHyperlink> _hyperlinks = new Dictionary<IXLAddress, XLHyperlink>();

        public IEnumerator<XLHyperlink> GetEnumerator()
        {
            return _hyperlinks.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(XLHyperlink hyperlink)
        {
            _hyperlinks.Add(hyperlink.Cell.Address, hyperlink);
        }

        public void Delete(XLHyperlink hyperlink)
        {
            _hyperlinks.Remove(hyperlink.Cell.Address);
        }

        public void Delete(IXLAddress address)
        {
            _hyperlinks.Remove(address);
        }

        public bool TryDelete(IXLAddress address)
        {
            if (_hyperlinks.ContainsKey(address))
            {
                _hyperlinks.Remove(address);
                return true;
            }

            return false;
        }

        public XLHyperlink Get(IXLAddress address)
        {
            return _hyperlinks[address];
        }

        public bool TryGet(IXLAddress address, out XLHyperlink hyperlink)
        {
            return _hyperlinks.TryGetValue(address, out hyperlink);
        }
    }
}
