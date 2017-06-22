using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLHyperlinks: IXLHyperlinks
    {
        private Dictionary<IXLAddress, XLHyperlink> hyperlinks = new Dictionary<IXLAddress, XLHyperlink>();
        public IEnumerator<XLHyperlink> GetEnumerator()
        {
            return hyperlinks.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(XLHyperlink hyperlink)
        {
            hyperlinks.Add(hyperlink.Cell.Address, hyperlink);
        }

        public void Delete(XLHyperlink hyperlink)
        {
            hyperlinks.Remove(hyperlink.Cell.Address);
        }

        public void Delete(IXLAddress address)
        {
            hyperlinks.Remove(address);
        }

    }
}
