// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using ClosedXML.Attributes;
using ClosedXML.Excel;

using workbook.Excel;

namespace ClosedXML.Extensions;

public static class ExportExtensions
{
    /// <summary>
    /// Easily exports data to an Excel workbook.
    /// </summary>
    /// <typeparam name="TSource">The type of data to be exported, must implement <see cref="IXLExportable"/>.</typeparam>
    /// <param name="data">The data to be exported.</param>
    /// <param name="options">Optional export options.</param>
    /// <param name="columnCallback">Optional callback function for customizing column drawing.</param>
    /// <param name="rowCallback">Optional callback function for customizing row drawing.</param>
    /// <returns>An Excel workbook containing the exported data.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="data"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="data"/> contains an empty column.</exception>
    public static XLWorkbook EasyExport<TSource>(
        this IEnumerable<IXLExportable> data,
        XLExportOptions? options = null,
        Func<XLExportField, XLExportDrawCellResult>? columnCallback = null,
        Func<XLExportField, TSource, int, XLExportDrawCellResult>? rowCallback = null)
        where TSource : IXLExportable
    {
        if (data is null)
            throw new ArgumentNullException(nameof(data), "Data cannot be null.");

        var workbook = new XLWorkbook();

        var type = typeof(TSource);
        var columns = GetColumns(type);

        if (columns.Count == 0)
            throw new ArgumentException("The 'Data' argument contains an empty column.", nameof(data));

        var sheetName = string.IsNullOrEmpty(options?.SheetName) ? "Sheet1" : options.SheetName;
        var worksheet = workbook.Worksheets.Add(sheetName);

        int columnIndex = 0,
            rowIndex = 2,
            columnLength = columns.Count,
            rowLength = data.Count();

        var maxColumnCode = GetColumnCodeByIndex(columnLength);

        #region Column Handling

        worksheet.Row(1).Height = 25;
        worksheet.Range($"A1:{maxColumnCode}1")
            .Style.Font.SetBold(options?.Column?.Bold ?? true)
            .Border.SetInsideBorder(XLBorderStyleValues.Thin)
            .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
            .Fill.SetBackgroundColor(options?.Column?.BackgroundColor ?? XLColor.FromArgb(253, 233, 217))
            .Font.SetFontColor(options?.Column?.TextColor ?? XLColor.Black)
            .Alignment.SetHorizontal(options?.Column?.TextAlignHorizontal ?? XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(options?.Column?.TextAlignVertical ?? XLAlignmentVerticalValues.Center)
            .Font.SetFontSize(options?.Column?.FontSize ?? 12);

        foreach (var columnProperty in columns)
        {
            if (columnProperty.Property == null) continue;
            var attribute = columnProperty.Property.GetCustomAttribute<XLColumnAttribute>();
            columnCallback?.Invoke(columnProperty);

            worksheet.Cell(1, ++columnIndex).Value = attribute.Header ?? columnProperty.Property.Name;
        }

        #endregion Column Handling

        #region Row Handling

        if (rowLength > 0)
        {
            var rowRange = $"A2:{maxColumnCode}{rowLength + 1}";
            worksheet.Rows(2, rowLength).Height = 20;

            worksheet.Range(rowRange)
                .Style.Border.SetInsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Fill.SetBackgroundColor(options?.Row?.BackgroundColor ?? XLColor.White)
                .Font.SetFontColor(options?.Row?.TextColor ?? XLColor.Black)
                .Alignment.SetHorizontal(options?.Row?.TextAlignHorizontal ?? XLAlignmentHorizontalValues.Center)
                .Alignment.SetVertical(options?.Row?.TextAlignVertical ?? XLAlignmentVerticalValues.Center)
                .Font.SetFontSize(options?.Row?.FontSize ?? 11)
                .Font.SetBold(options?.Row?.Bold ?? false);

            // Add striped rows
            if (!options?.RemoveRowStriped ?? true)
            {
                for (var i = 2; i <= rowLength + 1; i++)
                {
                    if (i % 2 != 0)
                    {
                        worksheet.Range($"A{i}:{maxColumnCode}{i}").Style.Fill.SetBackgroundColor(XLColor.FromArgb(222, 226, 230));
                    }
                }
            }

            foreach (var row in data)
            {
                for (var i = 0; i < columnIndex; i++)
                {
                    var columnProperty = columns[i];
                    var callbackResult = rowCallback?.Invoke(columnProperty, (TSource)row, rowIndex);
                    callbackResult ??= new XLExportDrawCellResult
                    {
                        Value = GetPropertyValue(row, columnProperty)
                    };

                    if (callbackResult.IsSkip) continue;

                    var cell = worksheet.Cell(rowIndex, i + 1);
                    cell.Value = XLCellValue.FromObject(callbackResult.Value);

                    if (callbackResult.Options is null) continue;

                    var callbackResultOptions = callbackResult.Options;
                    cell.Style.Font.SetFontColor(callbackResultOptions?.TextColor ?? XLColor.Black)
                        .Fill.SetBackgroundColor(callbackResultOptions?.BackgroundColor ?? XLColor.White)
                        .Font.SetFontSize(callbackResultOptions?.FontSize ?? 11)
                        .Font.SetBold(callbackResultOptions?.Bold ?? false);
                }

                rowIndex++;
            }
        }

        #endregion Row Handling

        // Global Format
        worksheet.Columns($"A:{maxColumnCode}").AdjustToContents();

        if (workbook.Worksheets.Count == 0)
        {
            workbook.AddWorksheet();
        }

        return workbook;
    }

    private static string GetColumnCodeByIndex(int index)
    {
        var columnName = string.Empty;

        while (index > 0)
        {
            var modulo = (index - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            index = (index - modulo) / 26;
        }

        return columnName;
    }

    private static List<XLExportField> GetColumns(Type type, string? parentName = null)
    {
        var columns = type.GetProperties()
            .Where(ShouldIncludePropertyWithXLColumnAttribute)
            .OrderBy(GetXLColumnAttributeOrder)
            .ToList();

        var result = new List<XLExportField>();
        foreach (var column in columns)
        {
            var isRootType = column.PropertyType.IsBasicType();
            if (isRootType)
            {
                result.Add(new XLExportField(column, parentName));
            }
            else
            {
                var nestedColumns = GetColumns(column.PropertyType, column.Name);
                result.AddRange(nestedColumns);
            }
        }

        return result;
    }

    private static object? GetPropertyValue(IXLExportable row, XLExportField field)
    {
        if (field.Property is null)
            return null;

        try
        {
            if (string.IsNullOrEmpty(field.ParentName))
            {
                return field.Property.GetValue(row);
            }

            var parentProperty = row.GetType().GetProperty(field.ParentName);
            var parentValue = parentProperty.GetValue(row);

            return field.Property.GetValue(parentValue);
        }
        catch
        {
            return null;
        }
    }

    private static int GetXLColumnAttributeOrder(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<XLColumnAttribute>();
        return attribute is null || attribute.Order < 1 ? int.MaxValue : attribute.Order;
    }

    private static bool ShouldIncludePropertyWithXLColumnAttribute(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<XLColumnAttribute>();
        return attribute is { Ignore: false };
    }
}
