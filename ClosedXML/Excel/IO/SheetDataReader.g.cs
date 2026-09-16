#nullable enable

using System.Collections.Generic;
using ClosedXML.IO;
using StringItem = ClosedXML.Excel.CalcEngine.OneOf<string, ClosedXML.Excel.XLImmutableRichText>;

namespace ClosedXML.Excel.IO;

internal partial class SheetDataReader
{
    private Xpr ParseSheetData(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        OnSheetDataParsing();

        while (ParseRow("row", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'row' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnSheetDataParsed();
        return Xpr.Success();
    }

    partial void OnSheetDataParsing();

    partial void OnSheetDataParsed();

    private Xpr ParseRow(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var r = _reader.GetOptionalUInt("r");
        var spans = _reader.GetOptionalString("spans");
        var s = _reader.GetOptionalUInt("s") ?? 0;
        var customFormat = _reader.GetOptionalBool("customFormat") ?? false;
        var ht = _reader.GetOptionalDouble("ht");
        var hidden = _reader.GetOptionalBool("hidden") ?? false;
        var customHeight = _reader.GetOptionalBool("customHeight") ?? false;
        var outlineLevel = _reader.GetOptionalUByte("outlineLevel") ?? 0;
        var collapsed = _reader.GetOptionalBool("collapsed") ?? false;
        var thickTop = _reader.GetOptionalBool("thickTop") ?? false;
        var thickBot = _reader.GetOptionalBool("thickBot") ?? false;
        var ph = _reader.GetOptionalBool("ph") ?? false;
        var dyDescent = _reader.GetOptionalDouble("dyDescent", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

        OnRowParsing(r, spans, s, customFormat, ht, hidden, customHeight, outlineLevel, collapsed, thickTop, thickBot, ph, dyDescent);

        while (ParseCell("c", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'c' with cardinality 0-2147483647
        }
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        OnRowParsed(r, spans, s, customFormat, ht, hidden, customHeight, outlineLevel, collapsed, thickTop, thickBot, ph, dyDescent);
        return Xpr.Success();
    }

    partial void OnRowParsing(uint? r, string? spans, uint s, bool customFormat, double? ht, bool hidden, bool customHeight, byte outlineLevel, bool collapsed, bool thickTop, bool thickBot, bool ph, double? dyDescent);

    partial void OnRowParsed(uint? r, string? spans, uint s, bool customFormat, double? ht, bool hidden, bool customHeight, byte outlineLevel, bool collapsed, bool thickTop, bool thickBot, bool ph, double? dyDescent);

    private Xpr ParseCell(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var r = _reader.GetOptionalPoint("r");
        var s = _reader.GetOptionalUInt("s") ?? 0;
        var t = _reader.GetOptionalString("t") ?? "n";
        var cm = _reader.GetOptionalUInt("cm") ?? 0;
        var vm = _reader.GetOptionalUInt("vm") ?? 0;
        var ph = _reader.GetOptionalBool("ph") ?? false;

        OnCellParsing(r, s, t, cm, vm, ph);

        if (ParseCellFormula("f", _ns) is { IsSuccess: true })
        {
            // Optional element 'f' was present
        }
        var vResult = _reader.ParseXString("v", _ns);
        var v = vResult.IsSuccess ? vResult.Value : default(string?);
        var isResult = _rstReader.ParseCtRst("is", _ns);
        var @is = isResult.IsSuccess ? isResult.Value : default(StringItem?);
        if (ParseExtensionList("extLst", _ns) is { IsSuccess: true })
        {
            // Optional element 'extLst' was present
        }
        _reader.Close(elementName, ns);

        OnCellParsed(v, @is, r, s, t, cm, vm, ph);
        return Xpr.Success();
    }

    partial void OnCellParsing(Point? r, uint s, string t, uint cm, uint vm, bool ph);

    partial void OnCellParsed(string? v, StringItem? @is, Point? r, uint s, string t, uint cm, uint vm, bool ph);

    private Xpr ParseCellFormula(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var t = _reader.GetOptionalString("t") ?? string.Empty;
        var aca = _reader.GetOptionalBool("aca") ?? false;
        var @ref = _reader.GetOptionalArea("ref");
        var dt2D = _reader.GetOptionalBool("dt2D") ?? false;
        var dtr = _reader.GetOptionalBool("dtr") ?? false;
        var del1 = _reader.GetOptionalBool("del1") ?? false;
        var del2 = _reader.GetOptionalBool("del2") ?? false;
        var r1 = _reader.GetOptionalPoint("r1");
        var r2 = _reader.GetOptionalPoint("r2");
        var ca = _reader.GetOptionalBool("ca") ?? false;
        var si = _reader.GetOptionalUInt("si");
        var bx = _reader.GetOptionalBool("bx") ?? false;

        OnCellFormulaParsing(t, aca, @ref, dt2D, dtr, del1, del2, r1, r2, ca, si, bx);

        var formula = _reader.GetContent();
        _reader.Close(elementName, ns);

        OnCellFormulaParsed(formula, t, aca, @ref, dt2D, dtr, del1, del2, r1, r2, ca, si, bx);
        return Xpr.Success();
    }

    partial void OnCellFormulaParsing(string t, bool aca, Area? @ref, bool dt2D, bool dtr, bool del1, bool del2, Point? r1, Point? r2, bool ca, uint? si, bool bx);

    partial void OnCellFormulaParsed(string formula, string t, bool aca, Area? @ref, bool dt2D, bool dtr, bool del1, bool del2, Point? r1, Point? r2, bool ca, uint? si, bool bx);
}
