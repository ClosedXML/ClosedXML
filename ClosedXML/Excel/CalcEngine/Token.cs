namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Represents a node in the expression tree.
    /// </summary>
    internal class Token
    {
        // ** fields
        public TKID ID;

        public TKTYPE Type;
        public object Value;

        // ** ctor
        public Token(object value, TKID id, TKTYPE type)
        {
            Value = value;
            ID = id;
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

    /// <summary>
    /// Token ID (used when evaluating expressions)
    /// </summary>
    internal enum TKID
    {
        GT, LT, GE, LE, EQ, NE, // COMPARE
        ADD, SUB, // ADDSUB
        MUL, DIV, DIVINT, MOD, // MULDIV
        POWER, // POWER
        DIV100, // MULTIV_UNARY
        OPEN, CLOSE, END, COMMA, PERIOD, // GROUP
        ATOM, // LITERAL, IDENTIFIER
        CONCAT
    }
}
