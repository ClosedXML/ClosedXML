using DocumentFormat.OpenXml.Spreadsheet;
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
    private readonly Dictionary<int, string> _numberFormats = new();

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

    internal void LoadNumberFormats(NumberingFormats? numberingFormats)
    {
        if (numberingFormats is null)
            return;

        foreach (var nf in numberingFormats.ChildElements.Cast<NumberingFormat>())
        {
            var numberFormatId = checked((int?)nf.NumberFormatId?.Value);
            var formatCode = nf.FormatCode?.Value;
            if (numberFormatId is null || string.IsNullOrEmpty(formatCode))
                continue;

            _numberFormats.Add(numberFormatId.Value, formatCode);
        }
    }

    internal XLNumberFormatValue? GetNumberFormat(int? numberFormatId)
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
                NumberFormatId = numberFormatId.Value,
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
