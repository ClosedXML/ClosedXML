using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests;

internal static class RangeExtensions
{
    public static string ToSpaceList(this IEnumerable<IXLRange> ranges)
    {
        return string.Join(" ", ranges.Select(r => r.RangeAddress.ToString(XLReferenceStyle.A1, false)));
    }
}
