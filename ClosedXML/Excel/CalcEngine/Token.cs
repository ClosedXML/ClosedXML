using System.Diagnostics;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Represents a node in the expression tree.
    /// </summary>
    [DebuggerDisplay("{Value} ({ID} {Type})")]
    internal class Token
    {
        // ** fields
        public TKTYPE Type;
        public object Value;

        // ** ctor
        public Token(object value, TKTYPE type)
        {
            Value = value;
            Type = type;
        }
    }

    /// <summary>
    /// Token types (used when building expressions, sequence defines operator priority)
    /// </summary>
    internal enum TKTYPE
    {
        COMPARE,     // < > = <= >=
        ADDSUB,      // + -
        MULDIV,      // * /
        POWER,       // ^
        MULDIV_UNARY,// %
        GROUP,       // ( ) , .
        LITERAL,     // 123.32, "Hello", etc.
        IDENTIFIER,  // functions, external objects, bindings
        ERROR        // e.g. #REF!
    }
}
