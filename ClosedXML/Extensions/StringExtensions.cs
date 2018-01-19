using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXML.Extensions
{
    internal static class StringExtensions
    {
        internal static string EscapeSheetName(this String sheetName)
        {
            if (sheetName.Contains("'") ||
                sheetName.Contains(" "))
                return string.Format("'{0}'", sheetName.Replace("'", "''"));

            return sheetName;
        }

        internal static string UnescapeSheetName(this String sheetName)
        {
            return sheetName
                .Trim('\'')
                .Replace("''", "'");
        }

        internal static String HashPassword(this String password)
        {
            if (password == null) return null;

            Int32 pLength = password.Length;
            Int32 hash = 0;
            if (pLength == 0) return String.Empty;

            for (Int32 i = pLength - 1; i >= 0; i--)
            {
                hash ^= password[i];
                hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
            }
            hash ^= 0x8000 | 'N' << 8 | 'K';
            hash ^= pLength;
            return hash.ToString("X");
        }

    }
}
