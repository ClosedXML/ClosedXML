using System.Collections.Generic;

namespace ClosedXML.Excel
{
    using System;

    internal class XLStylizedContainer: IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        readonly IXLStylized _container;
        public XLStylizedContainer(IXLStyle style, IXLStylized container)
        {
            Style = style;
            _container = container;
            RangesUsed = container.RangesUsed;
        }

        public IXLStyle Style { get; set; }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                var _backupUpdatingStyle = _container.UpdatingStyle;
                _container.UpdatingStyle = true;
                yield return Style;
                _container.UpdatingStyle = _backupUpdatingStyle;
            }
        }

        public bool UpdatingStyle { get; set; }

        public IXLStyle InnerStyle { get; set; }

        public IXLRanges RangesUsed { get; set; }
    }
}
