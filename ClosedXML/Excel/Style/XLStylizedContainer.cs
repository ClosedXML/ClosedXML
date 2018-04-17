using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLStylizedContainer : XLStylizedBase, IXLStylized
    {
        protected readonly IXLStylized _container;

        public XLStylizedContainer(IXLStyle style, IXLStylized container) : base((style as XLStyle).Value)
        {
            _container = container;
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
            }
        }

        public override IXLRanges RangesUsed
        {
            get { return _container.RangesUsed; }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                if (_container is XLStylizedBase)
                    yield return _container as XLStylizedBase;

                yield break;
            }
        }
    }
}
