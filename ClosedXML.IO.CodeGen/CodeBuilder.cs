using System.Text;

namespace ClosedXML.IO.CodeGen;

internal class CodeBuilder
{
    private readonly StringBuilder _sb;
    private int _indentLevel;
    private bool _methodWritten;

    public CodeBuilder(StringBuilder sb)
    {
        _sb = sb;
    }

    internal CodeBuilder AddLine(string s)
    {
        AddIndentedLine(s);
        return this;
    }

    internal CodeBuilder OpenBrace()
    {
        AddIndentedLine("{");
        _indentLevel++;
        return this;
    }

    internal CodeBuilder CloseBrace()
    {
        _indentLevel--;
        AddIndentedLine("}");
        return this;
    }

    internal CodeBuilder StartMethod(string signature)
    {
        if (_methodWritten)
            _sb.AppendLine();

        AddLine(signature);
        _methodWritten = true;
        return this;
    }

    private void AddIndentedLine(string text)
    {
        for (var i = 0; i < _indentLevel; i++)
            _sb.Append("    ");

        _sb.AppendLine(text);
    }

    public override string ToString()
    {
        return _sb.ToString();
    }
}
