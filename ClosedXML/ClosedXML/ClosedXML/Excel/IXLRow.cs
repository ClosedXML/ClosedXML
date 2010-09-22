using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRow: IXLRange
    {
        Double Height { get; set; }
        void Delete();
        
    }

    public static class IXLRowMethods
    {

    }
}
