﻿using System;

namespace ClosedXML.Excel
{
    internal class XLCFIconSet : IXLCFIconSet
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFIconSet(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }
        public IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, String value, XLCFContentType type)
        {
            _conditionalFormat.IconSetOperators.Add(setOperator);
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
            _conditionalFormat.ContentTypes.Add(type);
            return new XLCFIconSet(_conditionalFormat);
        }
        public IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, Double value, XLCFContentType type)
        {
            return AddValue(setOperator, value.ToInvariantString(), type);
        }
    }
}
