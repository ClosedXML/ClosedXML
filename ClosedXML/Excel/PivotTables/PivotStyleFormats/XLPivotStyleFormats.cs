// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPivotStyleFormats : IXLPivotStyleFormats
    {
        private readonly IXLPivotField _pivotField;
        private readonly Dictionary<XLPivotStyleFormatElement, IXLPivotStyleFormat> _styleFormats = new Dictionary<XLPivotStyleFormatElement, IXLPivotStyleFormat>();

        public XLPivotStyleFormats()
            : this(null)
        { }

        public XLPivotStyleFormats(IXLPivotField pivotField)
        {
            this._pivotField = pivotField;
        }

        #region IXLPivotStyleFormats members

        public IXLPivotStyleFormat ForElement(XLPivotStyleFormatElement element)
        {
            if (element == XLPivotStyleFormatElement.None)
                throw new ArgumentException("Choose an enum value that represents an element", nameof(element));

            if (!_styleFormats.ContainsKey(element))
                _styleFormats.Add(element, new XLPivotStyleFormat(_pivotField) { AppliesTo = element });

            return _styleFormats[element];
        }

        public IEnumerator<IXLPivotStyleFormat> GetEnumerator()
        {
            return _styleFormats.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion IXLPivotStyleFormats members

        public void Add(IXLPivotStyleFormat styleFormat)
        {
            _styleFormats.Add(styleFormat.AppliesTo, styleFormat);
        }

        public void AddRange(IEnumerable<IXLPivotStyleFormat> styleFormats)
        {
            foreach (var styleFormat in styleFormats)
                Add(styleFormat);
        }
    }
}
