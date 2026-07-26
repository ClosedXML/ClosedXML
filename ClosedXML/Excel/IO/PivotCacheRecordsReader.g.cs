#nullable enable

using System.Collections.Generic;
using ClosedXML.IO;

namespace ClosedXML.Excel.IO;

internal partial class PivotCacheRecordsReader
{
    private Xpr ParseRecord(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        // Choice with cardinality 1-n
        var choiceCount = 0;
        while (true)
        {
            if (ParseMissing("m", _ns) is { IsSuccess: true })
            {
                // Choice m was successfully parsed
            }
            else if (ParseNumber("n", _ns) is { IsSuccess: true })
            {
                // Choice n was successfully parsed
            }
            else if (ParseBoolean("b", _ns) is { IsSuccess: true })
            {
                // Choice b was successfully parsed
            }
            else if (ParseError("e", _ns) is { IsSuccess: true })
            {
                // Choice e was successfully parsed
            }
            else if (ParseString("s", _ns) is { IsSuccess: true })
            {
                // Choice s was successfully parsed
            }
            else if (ParseDateTime("d", _ns) is { IsSuccess: true })
            {
                // Choice d was successfully parsed
            }
            else if (ParseIndex("x", _ns) is { IsSuccess: true })
            {
                // Choice x was successfully parsed
            }
            else
            {
                break;
            }
            choiceCount++;
        }
        if(choiceCount == 0)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnRecordParsed();
        return Xpr.Success();
    }

    partial void OnRecordParsed();

    private Xpr ParseMissing(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

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

        while (ParseTuples("tpls", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tpls' with cardinality 0-2147483647
        }
        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnMissingParsed(u, f, c, cp, @in, bc, fc, i, un, st, b);
        return Xpr.Success();
    }

    partial void OnMissingParsed(bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private Xpr ParseNumber(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

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

        while (ParseTuples("tpls", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tpls' with cardinality 0-2147483647
        }
        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnNumberParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
        return Xpr.Success();
    }

    partial void OnNumberParsed(double v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private Xpr ParseBoolean(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var v = _reader.GetBool("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");

        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnBooleanParsed(v, u, f, c, cp);
        return Xpr.Success();
    }

    partial void OnBooleanParsed(bool v, bool? u, bool? f, string? c, uint? cp);

    private Xpr ParseError(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

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

        if (ParseTuples("tpls", _ns) is { IsSuccess: true })
        {
            // Optional element 'tpls' was present
        }
        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnErrorParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
        return Xpr.Success();
    }

    partial void OnErrorParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private Xpr ParseString(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

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

        while (ParseTuples("tpls", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tpls' with cardinality 0-2147483647
        }
        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnStringParsed(v, u, f, c, cp, @in, bc, fc, i, un, st, b);
        return Xpr.Success();
    }

    partial void OnStringParsed(string v, bool? u, bool? f, string? c, uint? cp, uint? @in, uint? bc, uint? fc, bool i, bool un, bool st, bool b);

    private Xpr ParseDateTime(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var v = _reader.GetDateTime("v");
        var u = _reader.GetOptionalBool("u");
        var f = _reader.GetOptionalBool("f");
        var c = _reader.GetOptionalXString("c");
        var cp = _reader.GetOptionalUInt("cp");

        while (ParseX("x", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'x' with cardinality 0-2147483647
        }
        _reader.Close(elementName, ns);

        OnDateTimeParsed(v, u, f, c, cp);
        return Xpr.Success();
    }

    partial void OnDateTimeParsed(System.DateTime v, bool? u, bool? f, string? c, uint? cp);

    private Xpr ParseIndex(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var v = _reader.GetUInt("v");

        _reader.Close(elementName, ns);

        OnIndexParsed(v);
        return Xpr.Success();
    }

    partial void OnIndexParsed(uint v);

    private Xpr ParseX(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var v = _reader.GetOptionalInt("v") ?? 0;

        _reader.Close(elementName, ns);

        OnXParsed(v);
        return Xpr.Success();
    }

    partial void OnXParsed(int v);

    private Xpr ParseTuples(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var c = _reader.GetOptionalUInt("c");

        var tplCount = 0;
        while (ParseTuple("tpl", _ns) is { IsSuccess: true })
        {
            // Parsed another element 'tpl' with cardinality 1-2147483647
            tplCount++;
        }

        if (tplCount < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }
        _reader.Close(elementName, ns);

        OnTuplesParsed(c);
        return Xpr.Success();
    }

    partial void OnTuplesParsed(uint? c);

    private Xpr ParseTuple(string elementName, string ns)
    {
        if (!_reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        var fld = _reader.GetOptionalUInt("fld");
        var hier = _reader.GetOptionalUInt("hier");
        var item = _reader.GetUInt("item");

        _reader.Close(elementName, ns);

        OnTupleParsed(fld, hier, item);
        return Xpr.Success();
    }

    partial void OnTupleParsed(uint? fld, uint? hier, uint item);
}
