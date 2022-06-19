using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLCustomProperties: IEnumerable<IXLCustomProperty>
    {
        void Add(IXLCustomProperty customProperty);
        void Add<T>(string name, T value);
        void Delete(string name);
        IXLCustomProperty CustomProperty(string name);
        
    }
}
