using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ClosedXML.Excel.Formatting;
using ClosedXML.Extensions;
using ClosedXML.IO;
using ClosedXML.Parser;
using DocumentFormat.OpenXml.Spreadsheet;
using StringItem = ClosedXML.Excel.CalcEngine.OneOf<string, ClosedXML.Excel.XLImmutableRichText>;

namespace ClosedXML.Excel.IO;

/// <summary>
/// A state based reader for <c>CT_SheetData</c>.
/// </summary>
internal partial class SheetDataReader
{
    private static readonly string[] DateCellFormats =
    {
        "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
        "yyyy-MM-ddTHH:mm", "yyyy-MM-dd" // Formats accepted by Excel.
    };

    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly XmlTreeReader _reader;
    private readonly XLWorksheet _ws;
    private readonly XLWorkbookStyles _styles;
    private readonly RstReader _rstReader;
    private readonly List<StringItem> _sst;
    private readonly Dictionary<uint, string> _sharedFormulasR1C1 = new();

    private int _row;
    private int _column;
    private string? _formulaText;
    private string? _formulaType;
    private bool _formulaAca;
    private Area? _formulaArea;
    private bool _formulaDt2D;
    private bool _formulaDtr;
    private bool _formulaDel1;
    private bool _formulaDel2;
    private Point? _formulaR1;
    private Point? _formulaR2;
    private bool _formulaCa;
    private uint? _formulaSi;

    internal SheetDataReader(XmlTreeReader reader, XLWorksheet ws, XLWorkbookStyles styles, List<StringItem> sst)
    {
        _reader = reader;
        _ws = ws;
        _styles = styles;
        _rstReader = new RstReader(reader, styles);
        _sst = sst;
    }

    [MemberNotNullWhen(true, [nameof(_formulaText), nameof(_formulaType)])]
    private bool HasFormula { get; set; }

    internal Xpr ParseCtSheetData(string elementName, string ns)
    {
        _row = 0;
        _column = 0;
        ResetFormula();
        return ParseSheetData(elementName, ns);
    }

    private Xpr ParseExtensionList(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        _reader.Skip(elementName);
        return Xpr.Success();
    }

    partial void OnRowParsing(uint? r, string? spans, uint s, bool customFormat, double? ht, bool hidden, bool customHeight, byte outlineLevel, bool collapsed, bool thickTop, bool thickBot, bool ph, double? dyDescent)
    {
        // Row number is an optional attribute. If not specified, it should be a next row from the last read row.
        _row = r is null ? _row + 1 : checked((int)r.Value);
        _column = 1;
        _ = spans;
        _ = s;
        _ = customFormat;
        _ = ht;
        _ = hidden;
        _ = customHeight;
        _ = outlineLevel;
        _ = collapsed;
        _ = thickTop;
        _ = thickBot;
        _ = ph;
        _ = dyDescent;
    }

    partial void OnRowParsed(uint? r, string? spans, uint s, bool customFormat, double? ht, bool hidden, bool customHeight, byte outlineLevel, bool collapsed, bool thickTop, bool thickBot, bool ph, double? dyDescent)
    {
        var rowIndex = r is null ? _row : checked((int)r.Value);
        var xlRow = _ws.Row(rowIndex, false);

        _ = spans; // Spans is unreliable, ignore it.

        if (ht is not null)
        {
            xlRow.Height = ht.Value;
        }
        else
        {
            xlRow.Loading = true;
            xlRow.Height = _ws.RowHeight;
            xlRow.Loading = false;
        }

        if (dyDescent is not null)
            xlRow.DyDescent = dyDescent.Value;

        if (hidden)
            xlRow.Hide();

        if (collapsed)
            xlRow.Collapsed = true;

        if (outlineLevel > 0)
            xlRow.OutlineLevel = outlineLevel;

        if (ph)
            xlRow.ShowPhonetic = true;

        if (customFormat)
            xlRow.FormatValue = _styles.CellFormats[checked((int)s)];
    }

    partial void OnCellFormulaParsed(string formula, string t, bool aca, Area? @ref, bool dt2D, bool dtr, bool del1, bool del2, Point? r1, Point? r2, bool ca, uint? si, bool bx)
    {
        HasFormula = true;
        _formulaText = formula;
        _formulaType = t;
        _formulaAca = aca;
        _formulaArea = @ref;
        _formulaDt2D = dt2D;
        _formulaDtr = dtr;
        _formulaDel1 = del1;
        _formulaDel2 = del2;
        _formulaR1 = r1;
        _formulaR2 = r2;
        _formulaCa = ca;
        _formulaSi = si;
        _ = bx; // bx attribute of cell formula is not ever used, per MS-OI29500 2.1.620
    }

    partial void OnCellParsed(string? v, StringItem? @is, Point? r, uint s, string t, uint cm, uint vm, bool ph)
    {
        var cellAddress = r ?? new Point(_row, _column);

        var dataType = t switch
        {
            "b" => CellValues.Boolean,
            "n" => CellValues.Number,
            "e" => CellValues.Error,
            "s" => CellValues.SharedString,
            "str" => CellValues.String,
            "inlineStr" => CellValues.InlineString,
            "d" => CellValues.Date,
            _ => throw new FormatException("Unknown cell type.")
        };

        var xlCell = _ws.Cell(cellAddress.Row, cellAddress.Column);

        var xfId = checked((int)s);
        var cellFormat = _ws.Workbook.Styles.CellFormats[xfId];
        xlCell.FormatValue = cellFormat;

        var showPhonetic = ph;
        if (showPhonetic)
            xlCell.ShowPhonetic = true;

        var cellMetaIndex = cm != 0 ? cm : (uint?)null;
        if (cellMetaIndex is not null)
            xlCell.CellMetaIndex = cellMetaIndex.Value;

        var valueMetaIndex = vm != 0 ? vm : (uint?)null;
        if (valueMetaIndex is not null)
            xlCell.ValueMetaIndex = valueMetaIndex.Value;

        var cellHasFormula = HasFormula;
        XLCellFormula? formula = null;
        if (cellHasFormula)
            formula = SetCellFormula(cellAddress);

        // Unified code to load value. Value can be empty and only type specified (e.g. when formula doesn't save values)
        // String type is only for formulas, while shared string/inline string/date is only for pure cell values.
        if (v is not null)
        {
            SetCellValue(dataType, v, xlCell, cellFormat);
        }
        else
        {
            // A string cell must contain at least empty string.
            if (dataType.Equals(CellValues.SharedString) || dataType.Equals(CellValues.String))
                xlCell.SetOnlyValue(string.Empty);
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when value is missing. Formula can be null in some cases, e.g. slave cells of array formula.
        if (formula is not null && v is null)
        {
            formula.IsDirty = true;
        }

        // Inline text is dealt separately, because it is in a separate element.
        //var cellHasInlineString = @is is not null;
        if (@is is not null)
        {
            if (dataType == CellValues.InlineString)
            {
                xlCell.ShareString = false;
                if (@is.Value.TryPickT0(out var text, out var richText))
                    xlCell.SetOnlyValue(text.FixNewLines());
                else
                    SetCellText(xlCell, richText);
            }
        }

        if (_ws.Workbook.Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
        {
            // Internally ClosedXML stores cells as standard 1900-based style
            // so if a workbook is in 1904-format, we do that adjustment here and when saving.
            xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
        }

        _column = cellAddress.Column + 1;
        ResetFormula();
    }

    private XLCellFormula? SetCellFormula(Point cellAddress)
    {
        if (!HasFormula)
            return null;

        var formulaSlice = _ws.Internals.CellsCollection.FormulaSlice;
        var valueSlice = _ws.Internals.CellsCollection.ValueSlice;

        var formulaType = _formulaType switch
        {
            "" or "normal" => CellFormulaValues.Normal,
            "array" => CellFormulaValues.Array,
            "dataTable" => CellFormulaValues.DataTable,
            "shared" => CellFormulaValues.Shared,
            _ => throw new NotSupportedException("Unknown formula type.")
        };

        // Always set shareString flag to `false`, because the text result of
        // formula is stored directly in the sheet, not shared string table.
        XLCellFormula? formula = null;
        if (formulaType == CellFormulaValues.Normal)
        {
            formula = XLCellFormula.NormalA1(_formulaText);
            formulaSlice.Set(cellAddress, formula);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.Array && _formulaArea is { } arrayArea) // Child cells of an array may have array type, but not ref, that is reserved for master cell
        {
            var aca = _formulaAca;

            // Because cells are read from top-to-bottom, from left-to-right, none of child cells have
            // a formula yet. Also, Excel doesn't allow change of array data, only through parent formula.
            formula = XLCellFormula.Array(_formulaText, arrayArea, aca);
            formulaSlice.SetArray(arrayArea, formula);

            for (var col = arrayArea.FirstPoint.Column; col <= arrayArea.LastPoint.Column; ++col)
            {
                for (var row = arrayArea.FirstPoint.Row; row <= arrayArea.LastPoint.Row; ++row)
                {
                    valueSlice.SetShareString(cellAddress, false);
                }
            }
        }
        else if (formulaType == CellFormulaValues.Shared && _formulaSi is { } sharedIndex)
        {
            // Shared formulas are rather limited in use and parsing, even by Excel
            // https://stackoverflow.com/questions/54654993. Therefore we accept them,
            // but don't output them. Shared formula is created, when user in Excel
            // takes a supported formula and drags it to more cells.
            if (!_sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
            {
                // Spec: The first formula in a group of shared formulas is saved
                // in the f element. This is considered the 'master' formula cell.
                formula = XLCellFormula.NormalA1(_formulaText);
                formulaSlice.Set(cellAddress, formula);

                // The key reason why Excel hates shared formulas is likely relative addressing and the messy situation it creates
                var formulaR1C1 = FormulaConverter.ToR1C1(_formulaText, cellAddress.Row, cellAddress.Column);
                _sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
            }
            else
            {
                // Spec: The formula expression for a cell that is specified to be part of a shared formula
                // (and is not the master) shall be ignored, and the master formula shall override.
                var sharedFormulaA1 = FormulaConverter.ToA1(sharedR1C1Formula, cellAddress.Row, cellAddress.Column);
                formula = XLCellFormula.NormalA1(sharedFormulaA1);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.DataTable && _formulaArea is { } dataTableArea)
        {
            var is2D = _formulaDt2D;
            var input1Deleted = _formulaDel1;
            var input1 = _formulaR1 ?? throw PartStructureException.MissingAttribute("r1");
            if (is2D)
            {
                // Input 2 is only used for 2D tables
                var input2Deleted = _formulaDel2;
                var input2 = _formulaR2 ?? throw PartStructureException.MissingAttribute("r2");
                formula = XLCellFormula.DataTable2D(dataTableArea, input1, input1Deleted, input2, input2Deleted);
                formulaSlice.Set(cellAddress, formula);
            }
            else
            {
                var isRowDataTable = _formulaDtr;
                formula = XLCellFormula.DataTable1D(dataTableArea, input1, input1Deleted, isRowDataTable);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }

        return formula;
    }

    private void SetCellValue(CellValues dataType, string cellValue, XLCell xlCell, XLCellFormatValue format)
    {
        if (dataType == CellValues.Number)
        {
            // XLCell is by default blank, so no need to set it.
            if (double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var number))
            {
                var numberDataType = format.NumberFormat.GetNumberDataType();
                var cellNumber = numberDataType switch
                {
                    XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
                    XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
                    _ => number // Normal number
                };
                xlCell.SetOnlyValue(cellNumber);
            }
        }
        else if (dataType == CellValues.SharedString)
        {
            if (int.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var sharedStringId)
                && sharedStringId >= 0 && sharedStringId < _sst.Count)
            {
                var sharedString = _sst[sharedStringId];
                SetCellText(xlCell, sharedString);
            }
            else
                xlCell.SetOnlyValue(string.Empty);
        }
        else if (dataType == CellValues.String) // A plain string that is a result of a formula calculation
        {
            xlCell.SetOnlyValue(cellValue);
        }
        else if (dataType == CellValues.Boolean)
        {
            var isTrue = string.Equals(cellValue, "1", StringComparison.Ordinal) ||
                         string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
            xlCell.SetOnlyValue(isTrue);
        }
        else if (dataType == CellValues.Error)
        {
            if (XLErrorParser.TryParseError(cellValue, out var error))
                xlCell.SetOnlyValue(error);
        }
        else if (dataType == CellValues.Date)
        {
            // Technically, cell can contain date as ISO8601 string, but not rarely used due
            // to inconsistencies between ISO and serial date time representation.
            var date = DateTime.ParseExact(cellValue, DateCellFormats,
                XLHelper.ParseCulture,
                DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
            xlCell.SetOnlyValue(date);
        }
    }

    private void SetCellText(XLCell xlCell, StringItem sharedString)
    {
        var valueSlice = _ws.Internals.CellsCollection.ValueSlice;
        if (sharedString.TryPickT0(out var plainText, out var richText))
        {
            valueSlice.SetCellValue(xlCell.Point, plainText);
        }
        else
        {
            var cellFormat = _ws.GetStyleValue(xlCell.Point);
            var adjustedRichText = richText.WithBaseFont(cellFormat.Font);
            valueSlice.SetRichText(xlCell.Point, adjustedRichText);
        }
    }

    private void ResetFormula()
    {
        HasFormula = false;
        _formulaText = null;
        _formulaType = null;
        _formulaAca = false;
        _formulaArea = null;
        _formulaDt2D = false;
        _formulaDtr = false;
        _formulaDel1 = false;
        _formulaDel2 = false;
        _formulaR1 = null;
        _formulaR2 = null;
        _formulaCa = false;
        _formulaSi = null;
    }
}
