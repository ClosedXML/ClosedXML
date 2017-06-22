using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXML.Extensions
{
    internal static class StringExtensions
    {
        internal static string WrapSheetNameInQuotesIfRequired(this string sheetName)
        {
            if (sheetName.Contains(' '))
                return "'" + sheetName + "'";
            else
                return sheetName;
        }
    }
}
