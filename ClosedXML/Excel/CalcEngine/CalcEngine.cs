using ClosedXML.Excel.CalcEngine.Exceptions;
using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// CalcEngine parses strings and returns Expression objects that can
    /// be evaluated.
    /// </summary>
    /// <remarks>
    /// <para>This class has three extensibility points:</para>
    /// <para>Use the <b>DataContext</b> property to add an object's properties to the engine scope.</para>
    /// <para>Use the <b>RegisterFunction</b> method to define custom functions.</para>
    /// <para>Override the <b>GetExternalObject</b> method to add arbitrary variables to the engine scope.</para>
    /// </remarks>
    internal class CalcEngine
    {
        private const string defaultFunctionNameSpace = "_xlfn";

        //---------------------------------------------------------------------------

        #region ** fields

        // members
        private readonly FormulaParser _parser;
        private string _expr;                           // expression being parsed

        private int _len;                               // length of the expression being parsed
        private int _ptr;                               // current pointer into expression
        private char[] _idChars;                        // valid characters in identifiers (besides alpha and digits)
        private Token _currentToken;                    // current token being parsed
        private Token _nextToken;                       // next token being parsed. to be used by Peek
        private Dictionary<object, Token> _tkTbl;       // table with tokens (+, -, etc)
        private Dictionary<string, FunctionDefinition> _fnTbl;      // table with constants and functions (pi, sin, etc)
        private Dictionary<string, object> _vars;       // table with variables
        private object _dataContext;                    // object with properties
        private bool _optimize;                         // optimize expressions when parsing
        protected ExpressionCache _cache;               // cache with parsed expressions
        private CultureInfo _ci;                        // culture info used to parse numbers/dates
        private char _decimal, _listSep, _percent;      // localized decimal separator, list separator, percent sign

        #endregion ** fields

        //---------------------------------------------------------------------------

        #region ** ctor

        public CalcEngine()
        {
            CultureInfo = CultureInfo.InvariantCulture;
            _tkTbl = GetSymbolTable();
            _fnTbl = GetFunctionTable();
            _vars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            _cache = new ExpressionCache(this);
            _optimize = false;
            _parser = new FormulaParser(this, _fnTbl);
#if DEBUG
            //this.Test();
#endif
        }

        #endregion ** ctor

        //---------------------------------------------------------------------------

        #region ** object model

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Expression Parse(string expression)
        {
            // initialize
            _expr = expression;
            _len = _expr.Length;
            _ptr = 0;
            _currentToken = null;
            _nextToken = null;

            // skip leading equals sign
            if (_len > 0 && _expr[0] == '=')
                _ptr++;

            // skip leading +'s
            while (_len > _ptr && _expr[_ptr] == '+')
                _ptr++;

            // parse the expression
            var expr = _parser.ParseToAst(_expr);

            // optimize expression
            if (_optimize)
            {
                expr = expr.Optimize();
            }

            // done
            return expr;
        }

        /// <summary>
        /// Evaluates a string.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        public object Evaluate(string expression)
        {
            var x = _cache != null
                    ? _cache[expression]
                    : Parse(expression);
            return x.Evaluate();
        }

        /// <summary>
        /// Gets or sets whether the calc engine should keep a cache with parsed
        /// expressions.
        /// </summary>
        public bool CacheExpressions
        {
            get { return _cache != null; }
            set
            {
                if (value != CacheExpressions)
                {
                    _cache = value
                        ? new ExpressionCache(this)
                        : null;
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the calc engine should optimize expressions when
        /// they are parsed.
        /// </summary>
        public bool OptimizeExpressions
        {
            get { return _optimize; }
            set { _optimize = value; }
        }

        /// <summary>
        /// Gets or sets a string that specifies special characters that are valid for identifiers.
        /// </summary>
        /// <remarks>
        /// Identifiers must start with a letter or an underscore, which may be followed by
        /// additional letters, underscores, or digits. This string allows you to specify
        /// additional valid characters such as ':' or '!' (used in Excel range references
        /// for example).
        /// </remarks>
        public char[] IdentifierChars
        {
            get { return _idChars; }
            set { _idChars = value; }
        }

        /// <summary>
        /// Registers a function that can be evaluated by this <see cref="CalcEngine"/>.
        /// </summary>
        /// <param name="functionName">Function name.</param>
        /// <param name="parmMin">Minimum parameter count.</param>
        /// <param name="parmMax">Maximum parameter count.</param>
        /// <param name="fn">Delegate that evaluates the function.</param>
        public void RegisterFunction(string functionName, int parmMin, int parmMax, CalcEngineFunction fn)
        {
            _fnTbl.Add(functionName, new FunctionDefinition(parmMin, parmMax, fn));
        }

        /// <summary>
        /// Registers a function that can be evaluated by this <see cref="CalcEngine"/>.
        /// </summary>
        /// <param name="functionName">Function name.</param>
        /// <param name="parmCount">Parameter count.</param>
        /// <param name="fn">Delegate that evaluates the function.</param>
        public void RegisterFunction(string functionName, int parmCount, CalcEngineFunction fn)
        {
            RegisterFunction(functionName, parmCount, parmCount, fn);
        }

        /// <summary>
        /// Gets an external object based on an identifier.
        /// </summary>
        /// <remarks>
        /// This method is useful when the engine needs to create objects dynamically.
        /// For example, a spreadsheet calc engine would use this method to dynamically create cell
        /// range objects based on identifiers that cannot be enumerated at design time
        /// (such as "AB12", "A1:AB12", etc.)
        /// </remarks>
        public virtual object GetExternalObject(string identifier)
        {
            return null;
        }

        /// <summary>
        /// Gets or sets the DataContext for this <see cref="CalcEngine"/>.
        /// </summary>
        /// <remarks>
        /// Once a DataContext is set, all public properties of the object become available
        /// to the CalcEngine, including sub-properties such as "Address.Street". These may
        /// be used with expressions just like any other constant.
        /// </remarks>
        public virtual object DataContext
        {
            get { return _dataContext; }
            set { _dataContext = value; }
        }

        /// <summary>
        /// Gets the dictionary that contains function definitions.
        /// </summary>
        public Dictionary<string, FunctionDefinition> Functions
        {
            get { return _fnTbl; }
        }

        /// <summary>
        /// Gets the dictionary that contains simple variables (not in the DataContext).
        /// </summary>
        public Dictionary<string, object> Variables
        {
            get { return _vars; }
        }

        /// <summary>
        /// Gets or sets the <see cref="CultureInfo"/> to use when parsing numbers and dates.
        /// </summary>
        public CultureInfo CultureInfo
        {
            get { return _ci; }
            set
            {
                _ci = value;
                var nf = _ci.NumberFormat;
                _decimal = nf.NumberDecimalSeparator[0];
                _percent = nf.PercentSymbol[0];
                _listSep = _ci.TextInfo.ListSeparator[0];
            }
        }

        #endregion ** object model

        //---------------------------------------------------------------------------

        #region ** token/keyword tables

        private static readonly IDictionary<string, ErrorExpression.ExpressionErrorType> ErrorMap = new Dictionary<string, ErrorExpression.ExpressionErrorType>()
        {
            ["#REF!"] = ErrorExpression.ExpressionErrorType.CellReference,
            ["#VALUE!"] = ErrorExpression.ExpressionErrorType.CellValue,
            ["#DIV/0!"] = ErrorExpression.ExpressionErrorType.DivisionByZero,
            ["#NAME?"] = ErrorExpression.ExpressionErrorType.NameNotRecognized,
            ["#N/A"] = ErrorExpression.ExpressionErrorType.NoValueAvailable,
            ["#NULL!"] = ErrorExpression.ExpressionErrorType.NullValue,
            ["#NUM!"] = ErrorExpression.ExpressionErrorType.NumberInvalid
        };

        // build/get static token table
        private Dictionary<object, Token> GetSymbolTable()
        {
            if (_tkTbl == null)
            {
                _tkTbl = new Dictionary<object, Token>();
                AddToken('&', TKID.CONCAT, TKTYPE.ADDSUB);
                AddToken('+', TKID.ADD, TKTYPE.ADDSUB);
                AddToken('-', TKID.SUB, TKTYPE.ADDSUB);
                AddToken('(', TKID.OPEN, TKTYPE.GROUP);
                AddToken(')', TKID.CLOSE, TKTYPE.GROUP);
                AddToken('*', TKID.MUL, TKTYPE.MULDIV);
                AddToken('.', TKID.PERIOD, TKTYPE.GROUP);
                AddToken('/', TKID.DIV, TKTYPE.MULDIV);
                AddToken('\\', TKID.DIVINT, TKTYPE.MULDIV);
                AddToken('%', TKID.DIV100, TKTYPE.MULDIV_UNARY);
                AddToken('=', TKID.EQ, TKTYPE.COMPARE);
                AddToken('>', TKID.GT, TKTYPE.COMPARE);
                AddToken('<', TKID.LT, TKTYPE.COMPARE);
                AddToken('^', TKID.POWER, TKTYPE.POWER);
                AddToken("<>", TKID.NE, TKTYPE.COMPARE);
                AddToken(">=", TKID.GE, TKTYPE.COMPARE);
                AddToken("<=", TKID.LE, TKTYPE.COMPARE);

                // list separator is localized, not necessarily a comma
                // so it can't be on the static table
                //AddToken(',', TKID.COMMA, TKTYPE.GROUP);
            }
            return _tkTbl;
        }

        private void AddToken(object symbol, TKID id, TKTYPE type)
        {
            var token = new Token(symbol, id, type);
            _tkTbl.Add(symbol, token);
        }

        // build/get static keyword table
        private Dictionary<string, FunctionDefinition> GetFunctionTable()
        {
            if (_fnTbl == null)
            {
                // create table
                _fnTbl = new Dictionary<string, FunctionDefinition>(StringComparer.InvariantCultureIgnoreCase);

                // register built-in functions (and constants)
                Engineering.Register(this);
                Information.Register(this);
                Logical.Register(this);
                Lookup.Register(this);
                MathTrig.Register(this);
                Text.Register(this);
                Statistical.Register(this);
                DateAndTime.Register(this);
                Financial.Register(this);
            }
            return _fnTbl;
        }

        #endregion ** token/keyword tables

        //---------------------------------------------------------------------------

        #region ** private stuff

        private Expression ParseExpression()
        {
            GetToken();
            return ParseCompare();
        }

        private Expression ParseCompare()
        {
            var x = ParseAddSub();
            while (_currentToken.Type == TKTYPE.COMPARE)
            {
                var t = _currentToken;
                GetToken();
                var exprArg = ParseAddSub();
                x = new BinaryExpression(t, x, exprArg);
            }
            return x;
        }

        private Expression ParseAddSub()
        {
            var x = ParseMulDiv();
            while (_currentToken.Type == TKTYPE.ADDSUB)
            {
                var t = _currentToken;
                GetToken();
                var exprArg = ParseMulDiv();
                x = new BinaryExpression(t, x, exprArg);
            }
            return x;
        }

        private Expression ParseMulDiv()
        {
            var x = ParsePower();
            while (_currentToken.Type == TKTYPE.MULDIV)
            {
                var t = _currentToken;
                GetToken();
                var a = ParsePower();
                x = new BinaryExpression(t, x, a);
            }
            return x;
        }

        private Expression ParsePower()
        {
            var x = ParseMulDivUnary();
            while (_currentToken.Type == TKTYPE.POWER)
            {
                var t = _currentToken;
                GetToken();
                var a = ParseMulDivUnary();
                x = new BinaryExpression(t, x, a);
            }
            return x;
        }

        private Expression ParseMulDivUnary()
        {
            var x = ParseUnary();
            while (_currentToken.Type == TKTYPE.MULDIV_UNARY)
            {
                var t = _tkTbl['/'];
                var a = new Expression(100);
                x = new BinaryExpression(t, x, a);
                GetToken();
            }
            return x;
        }

        private Expression ParseUnary()
        {
            // unary plus and minus
            if (_currentToken.Type == TKTYPE.ADDSUB)
            {
                var sign = 1;
                do
                {
                    if (_currentToken.ID == TKID.SUB)
                        sign = -sign;
                    GetToken();
                } while (_currentToken.Type == TKTYPE.ADDSUB);
                var a = ParseAtom();
                var t = (sign == 1)
                    ? _tkTbl['+']
                    : _tkTbl['-'];
                return new UnaryExpression(t, a);
            }

            // not unary, return atom
            return ParseAtom();
        }

        private Expression ParseAtom()
        {
            string id;
            Expression x = null;

            switch (_currentToken.Type)
            {
                // literals
                case TKTYPE.LITERAL:
                    x = new Expression(_currentToken);
                    break;

                // identifiers
                case TKTYPE.IDENTIFIER:

                    // get identifier
                    id = (string)_currentToken.Value;

                    // Peek ahead to see whether we have a function name (which will be followed by parenthesis
                    // Or another identifier, like a named range or cell reference
                    if (PeekToken().ID == TKID.OPEN)
                    {
                        // look for functions
                        var foundFunction = _fnTbl.TryGetValue(id, out FunctionDefinition functionDefinition);
                        if (!foundFunction && id.StartsWith($"{defaultFunctionNameSpace}."))
                            foundFunction = _fnTbl.TryGetValue(id.Substring(defaultFunctionNameSpace.Length + 1), out functionDefinition);

                        if (!foundFunction)
                            throw new NameNotRecognizedException($"The identifier `{id}` was not recognised.");

                        var p = GetParameters();
                        var pCnt = p == null ? 0 : p.Count;
                        if (functionDefinition.ParmMin != -1 && pCnt < functionDefinition.ParmMin)
                        {
                            Throw(string.Format("Too few parameters for function '{0}'. Expected a minimum of {1} and a maximum of {2}.", id, functionDefinition.ParmMin, functionDefinition.ParmMax));
                        }
                        if (functionDefinition.ParmMax != -1 && pCnt > functionDefinition.ParmMax)
                        {
                            Throw(string.Format("Too many parameters for function '{0}'.Expected a minimum of {1} and a maximum of {2}.", id, functionDefinition.ParmMin, functionDefinition.ParmMax));
                        }
                        x = new FunctionExpression(functionDefinition, p);
                        break;
                    }

                    // look for simple variables (much faster than binding!)
                    if (_vars.ContainsKey(id))
                    {
                        x = new VariableExpression(_vars, id);
                        break;
                    }

                    // look for external objects
                    var xObj = GetExternalObject(id);
                    if (xObj == null)
                        throw new NameNotRecognizedException($"The identifier `{id}` was not recognised.");

                    x = new XObjectExpression(xObj);
                    break;

                // sub-expressions
                case TKTYPE.GROUP:

                    // Normally anything other than opening parenthesis is illegal here
                    // but Excel allows omitted parameters so return empty value expression.
                    if (_currentToken.ID != TKID.OPEN)
                    {
                        return new EmptyValueExpression();
                    }

                    // get expression
                    GetToken();
                    x = ParseCompare();

                    // check that the parenthesis was closed
                    if (_currentToken.ID != TKID.CLOSE)
                    {
                        Throw("Unbalanced parenthesis.");
                    }

                    break;

                case TKTYPE.ERROR:
                    x = new ErrorExpression((ErrorExpression.ExpressionErrorType)_currentToken.Value);
                    break;
            }

            // make sure we got something...
            if (x == null)
            {
                Throw();
            }

            // done
            GetToken();
            return x;
        }

        #endregion ** private stuff

        //---------------------------------------------------------------------------

        #region ** parser

        private static IDictionary<char, char> matchingClosingSymbols = new Dictionary<char, char>()
        {
            { '\'', '\'' },
            { '[',  ']' }
        };

        private Token ParseToken()
        {
            // eat white space
            while (_ptr < _len && _expr[_ptr] <= ' ')
            {
                _ptr++;
            }

            // are we done?
            if (_ptr >= _len)
            {
                return new Token(null, TKID.END, TKTYPE.GROUP);
            }

            // prepare to parse
            int i;
            var c = _expr[_ptr];

            // operators
            // this gets called a lot, so it's pretty optimized.
            // note that operators must start with non-letter/digit characters.
            var isLetter = char.IsLetter(c);
            var isDigit = char.IsDigit(c);

            var isEnclosed = matchingClosingSymbols.TryGetValue(c, out char matchingClosingSymbol);

            if (!isLetter && !isDigit && !isEnclosed)
            {
                // if this is a number starting with a decimal, don't parse as operator
                var nxt = _ptr + 1 < _len ? _expr[_ptr + 1] : '0';
                bool isNumber = c == _decimal && char.IsDigit(nxt);
                if (!isNumber)
                {
                    // look up localized list separator
                    if (c == _listSep)
                    {
                        _ptr++;
                        return new Token(c, TKID.COMMA, TKTYPE.GROUP);
                    }

                    // look up single-char tokens on table
                    if (_tkTbl.TryGetValue(c, out Token t))
                    {
                        // save token we found
                        var token = t;
                        _ptr++;

                        // look for double-char tokens (special case)
                        if (_ptr < _len
                            && (c == '>' || c == '<')
                            && _tkTbl.TryGetValue(_expr.Substring(_ptr - 1, 2), out t))
                        {
                            token = t;
                            _ptr++;
                        }

                        // found token on the table
                        return token;
                    }
                }
            }

            // parse numbers
            if (isDigit || c == _decimal)
            {
                var sci = false;
                var div = -1.0; // use double, not int (this may get really big)
                var val = 0.0;
                for (i = 0; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];

                    // digits always OK
                    if (char.IsDigit(c))
                    {
                        val = val * 10 + (c - '0');
                        if (div > -1)
                        {
                            div *= 10;
                        }
                        continue;
                    }

                    // one decimal is OK
                    if (c == _decimal && div < 0)
                    {
                        div = 1;
                        continue;
                    }

                    // scientific notation?
                    if ((c == 'E' || c == 'e') && !sci)
                    {
                        sci = true;
                        c = _expr[_ptr + i + 1];
                        if (c == '+' || c == '-') i++;
                        continue;
                    }

                    // end of literal
                    break;
                }

                // end of number, get value
                if (!sci)
                {
                    // much faster than ParseDouble
                    if (div > 1)
                    {
                        val /= div;
                    }
                }
                else
                {
                    var lit = _expr.Substring(_ptr, i);
                    val = ParseDouble(lit, _ci);
                }

                if (c != ':')
                {
                    // advance pointer and return
                    _ptr += i;

                    // build token
                    return new Token(val, TKID.ATOM, TKTYPE.LITERAL);
                }
            }

            // parse strings
            if (c == '\"')
            {
                // look for end quote, skip double quotes
                for (i = 1; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];
                    if (c != '\"') continue;
                    char cNext = i + _ptr < _len - 1 ? _expr[_ptr + i + 1] : ' ';
                    if (cNext != '\"') break;
                    i++;
                }

                // check that we got the end of the string
                if (c != '\"')
                {
                    Throw("Can't find final quote.");
                }

                // end of string
                var lit = _expr.Substring(_ptr + 1, i - 1);
                _ptr += i + 1;
                return new Token(lit.Replace("\"\"", "\""), TKID.ATOM, TKTYPE.LITERAL);
            }

            // parse #REF! (and other errors) in formula
            if (c == '#' && ErrorMap.Any(pair => _len > _ptr + pair.Key.Length && _expr.Substring(_ptr, pair.Key.Length).Equals(pair.Key, StringComparison.OrdinalIgnoreCase)))
            {
                var errorPair = ErrorMap.Single(pair => _len > _ptr + pair.Key.Length && _expr.Substring(_ptr, pair.Key.Length).Equals(pair.Key, StringComparison.OrdinalIgnoreCase));
                _ptr += errorPair.Key.Length;
                return new Token(errorPair.Value, TKID.ATOM, TKTYPE.ERROR);
            }

            // identifiers (functions, objects) must start with alpha or underscore
            if (!isEnclosed && !isLetter && c != '_' && (_idChars == null || !_idChars.Contains(c)))
            {
                Throw("Identifier expected.");
            }

            // and must contain only letters/digits/_idChars
            for (i = 1; i + _ptr < _len; i++)
            {
                c = _expr[_ptr + i];
                isLetter = char.IsLetter(c);
                isDigit = char.IsDigit(c);

                if (isEnclosed && c == matchingClosingSymbol)
                {
                    isEnclosed = false;
                    matchingClosingSymbol = '\0';

                    i++;
                    c = _expr[_ptr + i];
                    isLetter = char.IsLetter(c);
                    isDigit = char.IsDigit(c);
                }

                var disallowedSymbols = new List<char>() { '\\', '/', '*', '[', ':', '?' };
                if (isEnclosed && disallowedSymbols.Contains(c))
                    break;

                var allowedSymbols = new List<char>() { '_', '.' };

                if (!isLetter && !isDigit
                    && !(isEnclosed || allowedSymbols.Contains(c))
                    && (_idChars == null || !_idChars.Contains(c)))
                    break;
            }

            // got identifier
            var id = _expr.Substring(_ptr, i);
            _ptr += i;

            // If we have a true/false, return a literal
            if (bool.TryParse(id, out var b))
                return new Token(b, TKID.ATOM, TKTYPE.LITERAL);

            return new Token(id, TKID.ATOM, TKTYPE.IDENTIFIER);
        }

        private void GetToken()
        {
            if (_nextToken == null)
            {
                _currentToken = ParseToken();
            }
            else
            {
                _currentToken = _nextToken;
                _nextToken = null;
            }
        }

        private Token PeekToken() => _nextToken ??= ParseToken();

        private static double ParseDouble(string str, CultureInfo ci)
        {
            if (str.Length > 0 && str[str.Length - 1] == ci.NumberFormat.PercentSymbol[0])
            {
                str = str.Substring(0, str.Length - 1);
                return double.Parse(str, NumberStyles.Any, ci) / 100.0;
            }
            return double.Parse(str, NumberStyles.Any, ci);
        }

        private List<Expression> GetParameters() // e.g. myfun(a, b, c+2)
        {
            // check whether next token is a (,
            // restore state and bail if it's not
            var pos = _ptr;
            var tk = _currentToken;
            GetToken();
            if (_currentToken.ID != TKID.OPEN)
            {
                _ptr = pos;
                _currentToken = tk;
                return null;
            }

            // check for empty Parameter list
            pos = _ptr;
            GetToken();
            if (_currentToken.ID == TKID.CLOSE)
            {
                return null;
            }
            _ptr = pos;

            // get Parameters until we reach the end of the list
            var parms = new List<Expression>();
            var expr = ParseExpression();
            parms.Add(expr);
            while (_currentToken.ID == TKID.COMMA)
            {
                expr = ParseExpression();
                parms.Add(expr);
            }

            // make sure the list was closed correctly
            if (_currentToken.ID == TKID.OPEN)
                Throw("Unknown function: " + expr.LastParseItem);
            else if (_currentToken.ID != TKID.CLOSE)
                Throw("Syntax error: expected ')'");

            // done
            return parms;
        }

        private Token GetMember()
        {
            // check whether next token is a MEMBER token ('.'),
            // restore state and bail if it's not
            var pos = _ptr;
            var tk = _currentToken;
            GetToken();
            if (_currentToken.ID != TKID.PERIOD)
            {
                _ptr = pos;
                _currentToken = tk;
                return null;
            }

            // skip member token
            GetToken();
            if (_currentToken.Type != TKTYPE.IDENTIFIER)
            {
                Throw("Identifier expected");
            }
            return _currentToken;
        }

        #endregion ** parser

        //---------------------------------------------------------------------------

        #region ** static helpers

        private static void Throw()
        {
            Throw("Syntax error.");
        }

        private static void Throw(string msg)
        {
            throw new ExpressionParseException(msg);
        }

        #endregion ** static helpers
    }

    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    internal delegate object CalcEngineFunction(List<Expression> parms);
}
