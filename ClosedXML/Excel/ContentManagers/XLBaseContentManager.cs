using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.ContentManagers
{
    internal abstract class XLBaseContentManager
    {

    }

    internal abstract class XLBaseContentManager<T> : XLBaseContentManager
        where T : struct, IConvertible

    {
        protected readonly IDictionary<T, OpenXmlElement> contents = new Dictionary<T, OpenXmlElement>();

        public OpenXmlElement GetPreviousElementFor(T content)
        {
            var i = content.CastTo<int>();

            var previousElements = contents.Keys
                .Where(key => key.CastTo<int>() < i && contents[key] != null)
                .OrderBy(key => key.CastTo<int>());

            if (previousElements.Any())
                return contents[previousElements.Last()];
            else
                return null;
        }

        public void SetElement(T content, OpenXmlElement element)
        {
            contents[content] = element;
        }
    }
}
