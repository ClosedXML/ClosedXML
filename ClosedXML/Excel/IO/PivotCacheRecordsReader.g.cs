#nullable enable

using ClosedXML.IO;

namespace ClosedXML.Excel.IO;

internal partial class PivotCacheRecordsReader
{
    private void ParseRecord(string elementName)
    {
        do
        {
            if (_reader.TryOpen("m", _ns))
            {
                ParseMissing("m");
            }
            else if (_reader.TryOpen("n", _ns))
            {
                ParseNumber("n");
            }
            else if (_reader.TryOpen("b", _ns))
            {
                ParseBoolean("b");
            }
            else if (_reader.TryOpen("e", _ns))
            {
                ParseError("e");
            }
            else if (_reader.TryOpen("s", _ns))
            {
                ParseString("s");
            }
            else if (_reader.TryOpen("d", _ns))
            {
                ParseDateTime("d");
            }
            else if (_reader.TryOpen("x", _ns))
            {
                ParseIndex("x");
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(_reader);
            }
        }
        while (!_reader.TryClose(elementName, _ns));
        OnRecordParsed();
    }

    partial void OnRecordParsed();

    private void ParseMissing(string elementName)
    {
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        var @in = _reader.GetOptionalUInt("in");
        var bc = _reader.GetOptionalUIntHex("bc");
        var fc = _reader.GetOptionalUIntHex("fc");
        var i = _reader.GetOptionalBool("i") ?? false;
        var un = _reader.GetOptionalBool("un") ?? false;
        var st = _reader.GetOptionalBool("st") ?? false;
        var b = _reader.GetOptionalBool("b") ?? false;
        while (_reader.TryOpen("tpls", _ns))
        {
            ParseTuples("tpls");
        }
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnMissingParsed(u, f, c, cp, @in, bc, fc, i, un, st, b);
    }

    partial void OnMissingParsed(bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private void ParseNumber(string elementName)
    {
        var v = _reader.GetDouble("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        var @in = _reader.GetOptionalUInt("in");
        var bc = _reader.GetOptionalUIntHex("bc");
        var fc = _reader.GetOptionalUIntHex("fc");
        var i = _reader.GetOptionalBool("i") ?? false;
        var un = _reader.GetOptionalBool("un") ?? false;
        var st = _reader.GetOptionalBool("st") ?? false;
        var b = _reader.GetOptionalBool("b") ?? false;
        while (_reader.TryOpen("tpls", _ns))
        {
            ParseTuples("tpls");
        }
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnNumberParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
    }

    partial void OnNumberParsed(double v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private void ParseBoolean(string elementName)
    {
        var v = _reader.GetBool("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnBooleanParsed(v, u, f, c, cp);
    }

    partial void OnBooleanParsed(bool v, bool? u, bool? f, string? c, uint? cp);

    private void ParseError(string elementName)
    {
        var v = _reader.GetXString("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        var @in = _reader.GetOptionalUInt("in");
        var bc = _reader.GetOptionalUIntHex("bc");
        var fc = _reader.GetOptionalUIntHex("fc");
        var i = _reader.GetOptionalBool("i") ?? false;
        var un = _reader.GetOptionalBool("un") ?? false;
        var st = _reader.GetOptionalBool("st") ?? false;
        var b = _reader.GetOptionalBool("b") ?? false;
        if (_reader.TryOpen("tpls", _ns))
        {
            ParseTuples("tpls");
        }
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnErrorParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
    }

    partial void OnErrorParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private void ParseString(string elementName)
    {
        var v = _reader.GetXString("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        var @in = _reader.GetOptionalUInt("in");
        var bc = _reader.GetOptionalUIntHex("bc");
        var fc = _reader.GetOptionalUIntHex("fc");
        var i = _reader.GetOptionalBool("i") ?? false;
        var un = _reader.GetOptionalBool("un") ?? false;
        var st = _reader.GetOptionalBool("st") ?? false;
        var b = _reader.GetOptionalBool("b") ?? false;
        while (_reader.TryOpen("tpls", _ns))
        {
            ParseTuples("tpls");
        }
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnStringParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
    }

    partial void OnStringParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private void ParseDateTime(string elementName)
    {
        var v = _reader.GetDateTime("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");
        while (_reader.TryOpen("x", _ns))
        {
            ParseX("x");
        }
        _reader.Close(elementName, _ns);
        OnDateTimeParsed(v, u, f, c, cp);
    }

    partial void OnDateTimeParsed(System.DateTime v, bool? u, bool? f, string? c, uint? cp);

    private void ParseIndex(string elementName)
    {
        var v = _reader.GetUInt("v");
        _reader.Close(elementName, _ns);
        OnIndexParsed(v);
    }

    partial void OnIndexParsed(uint v);

    private void ParseX(string elementName)
    {
        var v = _reader.GetOptionalInt("v") ?? 0;
        _reader.Close(elementName, _ns);
        OnXParsed(v);
    }

    partial void OnXParsed(int v);

    private void ParseTuples(string elementName)
    {
        var c = _reader.GetOptionalUInt("c");
        _reader.Open("tpl", _ns);
        do
        {
            ParseTuple("tpl");
        }
        while (_reader.TryOpen("tpl", _ns));
        _reader.Close(elementName, _ns);
        OnTuplesParsed(c);
    }

    partial void OnTuplesParsed(uint? c);

    private void ParseTuple(string elementName)
    {
        var fld = _reader.GetOptionalUInt("fld");
        var hier = _reader.GetOptionalUInt("hier");
        var item = _reader.GetUInt("item");
        _reader.Close(elementName, _ns);
        OnTupleParsed(fld, hier, item);
    }

    partial void OnTupleParsed(uint? fld, uint? hier, uint item);
}
