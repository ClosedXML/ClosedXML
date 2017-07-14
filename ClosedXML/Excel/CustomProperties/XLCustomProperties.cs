using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCustomProperties: IXLCustomProperties, IEnumerable<IXLCustomProperty>
    {
        XLWorkbook workbook;
        public XLCustomProperties(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        private Dictionary<String, IXLCustomProperty> customProperties = new Dictionary<String, IXLCustomProperty>();
        public void Add(IXLCustomProperty customProperty)
        {
            customProperties.Add(customProperty.Name, customProperty);
        }
        public void Add<T>(String name, T value)
        {
            var cp = new XLCustomProperty(workbook) { Name = name, Value = value };
            Add(cp);
        }

        public void Delete(String name)
        {
            customProperties.Remove(name);
        }
        public IXLCustomProperty CustomProperty(String name)
        {
            return customProperties[name];
        }

        public IEnumerator<IXLCustomProperty> GetEnumerator()
        {
            return customProperties.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

      
    }
}
