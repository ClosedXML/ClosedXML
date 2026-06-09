using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests;

internal static class EnumerableExtensions
{
    public static string ToSpaceList(this IEnumerable<IXLRange> ranges, bool includeSheet = false)
    {
        return string.Join(" ", ranges.Select(r => r.RangeAddress.ToString(XLReferenceStyle.A1, includeSheet)));
    }

    public static string ToSpaceList(this IEnumerable<XLSheetRange> areas)
    {
        return string.Join(" ", areas);
    }
}
