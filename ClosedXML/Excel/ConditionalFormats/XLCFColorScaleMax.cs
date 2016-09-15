﻿using System;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMax : IXLCFColorScaleMax
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMax(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XLCFContentType type, String value, XLColor color)
        {
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(type);
        }
        public void Maximum(XLCFContentType type, Double value, XLColor color)
        {
            Maximum(type, value.ToInvariantString(), color);
        }
        public void HighestValue(XLColor color)
        {
            Maximum(XLCFContentType.Maximum, "0", color);
        }
    }
}
