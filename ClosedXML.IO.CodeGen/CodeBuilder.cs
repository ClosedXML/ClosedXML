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

    internal void AddLine(string s)
    {
        AddIndentedLine(s);
    }

    internal void OpenBrace()
    {
        AddIndentedLine("{");
        _indentLevel++;
    }

    internal void CloseBrace()
    {
        _indentLevel--;
        AddIndentedLine("}");
    }

    internal void StartMethod(string signature)
    {
        if (_methodWritten)
            _sb.AppendLine();

        AddLine(signature);
        _methodWritten = true;
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
