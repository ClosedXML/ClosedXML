using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.IO;

internal class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XLConditionalFormat>> _pivotCfs = new(XLHelper.SheetComparer);

    /// <summary>
    /// A dictionary of styles from <c>styles.xml</c>. Used in other places that reference number style by id reference.
    /// </summary>
    private readonly Dictionary<uint, string> _numberFormats = new();

    internal void AddPivotTableCf(string sheetName, XLConditionalFormat conditionalFormat)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
        {
            list = new List<XLConditionalFormat>();
            _pivotCfs[sheetName] = list;
        }

        list.Add(conditionalFormat);
    }

    internal XLConditionalFormat GetPivotCf(string sheetName, int priority)
    {
        if (!_pivotCfs.TryGetValue(sheetName, out var list))
            throw PivotCfNotFoundException(sheetName, priority);

        var pivotCf = list.SingleOrDefault(x => x.Priority == priority);
        if (pivotCf is null)
            throw PivotCfNotFoundException(sheetName, priority);

        return pivotCf;
    }

    internal void AddNumberFormat(uint numberFormatId, string numberFormat)
    {
        _numberFormats.Add(numberFormatId, numberFormat);
    }

    internal XLNumberFormatValue? GetNumberFormat(uint? numberFormatId)
    {
        if (numberFormatId is null)
        {
            return null;
        }

        if (_numberFormats.TryGetValue(numberFormatId.Value, out var formatCode))
        {
            var customFormatKey = new XLNumberFormatKey
            {
                NumberFormatId = -1,
                Format = formatCode,
            };
            return XLNumberFormatValue.FromKey(ref customFormatKey);
        }
        else
        {
            var predefinedFormatKey = new XLNumberFormatKey
            {
                NumberFormatId = checked((int)numberFormatId.Value),
                Format = string.Empty,
            };
            return XLNumberFormatValue.FromKey(ref predefinedFormatKey);
        }
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority)
    {
        return PartStructureException.ExpectedElementNotFound($"conditional formatting for pivot table in sheet {sheetName} with priority {priority}");
    }
}
