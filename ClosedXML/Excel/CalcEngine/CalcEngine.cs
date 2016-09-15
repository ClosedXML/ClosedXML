using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.CalcEngine.Functions;

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
		//---------------------------------------------------------------------------
		#region ** fields

		// members
		string		_expr;				        // expression being parsed
		int			_len;				        // length of the expression being parsed
		int			_ptr;				        // current pointer into expression
        string      _idChars;                   // valid characters in identifiers (besides alpha and digits)
		Token		_token;				        // current token being parsed
        Dictionary<object, Token> _tkTbl;       // table with tokens (+, -, etc)
        Dictionary<string, FunctionDefinition>  _fnTbl;      // table with constants and functions (pi, sin, etc)
        Dictionary<string, object> _vars;       // table with variables
        object _dataContext;                    // object with properties
        bool _optimize;                         // optimize expressions when parsing
        ExpressionCache _cache;                 // cache with parsed expressions
        CultureInfo _ci;                        // culture info used to parse numbers/dates
        char _decimal, _listSep, _percent;                // localized decimal separator, list separator, percent sign

        #endregion

        //---------------------------------------------------------------------------
        #region ** ctor

        public CalcEngine()
        {
            CultureInfo = CultureInfo.InvariantCulture;
            _tkTbl = GetSymbolTable();
            _fnTbl = GetFunctionTable();
            _vars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            _cache = new ExpressionCache(this);
            _optimize = true;
#if DEBUG
            //this.Test();
#endif
        }
        
        #endregion

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
			_len  = _expr.Length; 
			_ptr  = 0;

            // skip leading equals sign
            if (_len > 0 && _expr[0] == '=')
            {
                _ptr++;
            }

			// parse the expression
			var expr = ParseExpression();

			// check for errors
			if (_token.ID != TKID.END)
			{
                Throw();
			}

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
            var x = //Parse(expression);
                _cache != null
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
        public string IdentifierChars
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

		#endregion

        //---------------------------------------------------------------------------
        #region ** token/keyword tables

        // build/get static token table
        Dictionary<object, Token> GetSymbolTable()
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
        void AddToken(object symbol, TKID id, TKTYPE type)
        {
            var token = new Token(symbol, id, type);
            _tkTbl.Add(symbol, token);
        }

        // build/get static keyword table
        Dictionary<string, FunctionDefinition> GetFunctionTable()
        {
            if (_fnTbl == null)
            {
                // create table
                _fnTbl = new Dictionary<string, FunctionDefinition>(StringComparer.InvariantCultureIgnoreCase);

                // register built-in functions (and constants)
                Is.Register(this);
                Logical.Register(this);
                MathTrig.Register(this);
                Text.Register(this);
                Statistical.Register(this);
                DateAndTime.Register(this);
            }
            return _fnTbl;
        }

        #endregion

        //---------------------------------------------------------------------------
		#region ** private stuff

		Expression ParseExpression()
		{
			GetToken();
			return ParseCompare();
		}
		Expression ParseCompare()
		{
		    var x = ParseAddSub();
			while (_token.Type == TKTYPE.COMPARE)
			{
		        var t = _token;
				GetToken();
				var exprArg = ParseAddSub();
				x = new BinaryExpression(t, x, exprArg);
			}
			return x;
		}
		Expression ParseAddSub()
		{
			var x = ParseMulDiv();
			while (_token.Type == TKTYPE.ADDSUB)
			{
		        var t = _token;
				GetToken();
				var exprArg = ParseMulDiv();
				x = new BinaryExpression(t, x, exprArg);
			}
			return x;
        }
		Expression ParseMulDiv()
		{
			var x = ParsePower();
			while (_token.Type == TKTYPE.MULDIV)
			{
		        var t = _token;
				GetToken();
				var a = ParsePower();
				x = new BinaryExpression(t, x, a);
			}
			return x;
        }
		Expression ParsePower()
		{
			var x = ParseUnary();
		    while (_token.Type == TKTYPE.POWER)
			{
		        var t = _token;
				GetToken();
				var a = ParseUnary();
				x = new BinaryExpression(t, x, a);
			}
			return x;
		}
 		Expression ParseUnary()
		{ 
			// unary plus and minus
			if (_token.ID == TKID.ADD || _token.ID == TKID.SUB)
			{
				var t = _token;
		        GetToken();
                var a = ParseAtom();
                return new UnaryExpression(t, a);
			}

			// not unary, return atom
			return ParseAtom();
		}
		Expression ParseAtom()
		{
            string id;
            Expression x = null;
            FunctionDefinition fnDef = null;

			switch (_token.Type)
			{
				// literals
				case TKTYPE.LITERAL:
					x = new Expression(_token);
					break;

                // identifiers
                case TKTYPE.IDENTIFIER:

                    // get identifier
                    id = (string)_token.Value;

                    // look for functions
                    if (_fnTbl.TryGetValue(id, out fnDef))
                    {
                        var p = GetParameters();
                        var pCnt = p == null ? 0 : p.Count;
                        if (fnDef.ParmMin != -1 && pCnt < fnDef.ParmMin)
                        {
                            Throw("Too few parameters.");
                        }
                        if (fnDef.ParmMax != -1 && pCnt > fnDef.ParmMax)
                        {
                            Throw("Too many parameters.");
                        }
                        x = new FunctionExpression(fnDef, p);
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
                    if (xObj != null)
                    {
                        x = new XObjectExpression(xObj);
                        break;
                    }

                    // look for bindings
                    if (DataContext != null)
                    {
                        var list = new List<BindingInfo>();
                        for (var t = _token; t != null; t = GetMember())
                        {
                            list.Add(new BindingInfo((string)t.Value, GetParameters()));
                        }
                        x = new BindingExpression(this, list, _ci);
                        break;
                    }
                    Throw("Unexpected identifier");
                    break;

		        // sub-expressions
		        case TKTYPE.GROUP:

                    // anything other than opening parenthesis is illegal here
					if (_token.ID != TKID.OPEN)
					{
                        Throw("Expression expected.");
                    }

					// get expression
					GetToken();
					x = ParseCompare();

					// check that the parenthesis was closed
					if (_token.ID != TKID.CLOSE)
					{
						Throw("Unbalanced parenthesis.");
					}

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

		#endregion

		//---------------------------------------------------------------------------
		#region ** parser

        void GetToken()
        {
			// eat white space 
			while (_ptr < _len && _expr[_ptr] <= ' ')
			{
				_ptr++;
			}

			// are we done?
			if (_ptr >= _len)
			{
                _token = new Token(null, TKID.END, TKTYPE.GROUP);
				return;
			}

			// prepare to parse
            int i;
			var c = _expr[_ptr];

			// operators
			// this gets called a lot, so it's pretty optimized.
			// note that operators must start with non-letter/digit characters.
            var isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
            var isDigit = c >= '0' && c <= '9';
			if (!isLetter && !isDigit)
			{
				// if this is a number starting with a decimal, don't parse as operator
                var nxt = _ptr + 1 < _len ? _expr[_ptr + 1] : 0;
                bool isNumber = c == _decimal && nxt >= '0' && nxt <= '9';
				if (!isNumber)
				{
                    // look up localized list separator
                    if (c == _listSep)
                    {
                        _token = new Token(c, TKID.COMMA, TKTYPE.GROUP);
                        _ptr++;
                        return;
                    }
                    
                    // look up single-char tokens on table
                    Token tk;
                    if (_tkTbl.TryGetValue(c, out tk))
					{
						// save token we found
						_token = tk;
						_ptr++;

						// look for double-char tokens (special case)
						if (_ptr < _len && (c == '>' || c == '<'))
						{
                            if (_tkTbl.TryGetValue(_expr.Substring(_ptr - 1, 2), out tk))
							{
								_token = tk;
								_ptr++;
							}
						}

                        // found token on the table
						return;
					}
				}
			}

			// parse numbers
            if (isDigit || c == _decimal)
			{
				var sci = false;
                var pct = false;
                var div = -1.0; // use double, not int (this may get really big)
                var val = 0.0;
                for (i = 0; i + _ptr < _len; i++)
				{
					c = _expr[_ptr + i];

                    // digits always OK
                    if (c >= '0' && c <= '9')
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

                    // percentage?
                    if (c == _percent)
                    {
                        pct = true;
                        i++;
                        break;
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
                    if (pct)
                    {
                        val /= 100.0;
                    }
                }
                else
                {
                    var lit = _expr.Substring(_ptr, i);
                    val = ParseDouble(lit, _ci);
                }

                // build token
                _token = new Token(val, TKID.ATOM, TKTYPE.LITERAL);

                // advance pointer and return
                _ptr += i;
                return;
			}

			// parse strings
			if (c == '\"')
			{
				// look for end quote, skip double quotes
				for (i = 1; i + _ptr < _len; i++)
				{
					c = _expr[_ptr + i];
					if (c != '\"') continue;
					char cNext = i + _ptr < _len - 1 ? _expr[_ptr + i + 1]: ' ';
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
                _token = new Token(lit.Replace("\"\"", "\""), TKID.ATOM, TKTYPE.LITERAL);
				return;
			}

			// parse dates (review)
			if (c == '#')
			{
				// look for end #
				for (i = 1; i + _ptr < _len; i++)
				{
					c = _expr[_ptr + i];
					if (c == '#') break;
				}

				// check that we got the end of the date
				if (c != '#') 
				{
					Throw("Can't find final date delimiter ('#').");
				}

				// end of date
				var lit = _expr.Substring(_ptr + 1, i - 1);
				_ptr += i + 1;
                _token = new Token(DateTime.Parse(lit, _ci), TKID.ATOM, TKTYPE.LITERAL);
				return;
			}

            // identifiers (functions, objects) must start with alpha or underscore
            if (!isLetter && c != '_' && (_idChars == null || _idChars.IndexOf(c) < 0))
            {
                Throw("Identifier expected.");
            }

            // and must contain only letters/digits/_idChars
            for (i = 1; i + _ptr < _len; i++)
            {
                c = _expr[_ptr + i];
                isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
                isDigit = c >= '0' && c <= '9';
                if (!isLetter && !isDigit && c != '_' && (_idChars == null || _idChars.IndexOf(c) < 0))
                {
                    break;
                }
            }

            // got identifier
            var id = _expr.Substring(_ptr, i);
            _ptr += i;
            _token = new Token(id, TKID.ATOM, TKTYPE.IDENTIFIER);
		}
        static double ParseDouble(string str, CultureInfo ci)
        {
            if (str.Length > 0 && str[str.Length - 1] == ci.NumberFormat.PercentSymbol[0])
            {
                str = str.Substring(0, str.Length - 1);
                return double.Parse(str, NumberStyles.Any, ci) / 100.0;
            }
            return double.Parse(str, NumberStyles.Any, ci);
        }
        List<Expression> GetParameters() // e.g. myfun(a, b, c+2)
		{
			// check whether next token is a (, 
			// restore state and bail if it's not
			var pos  = _ptr;
			var tk = _token;
			GetToken();
			if (_token.ID != TKID.OPEN)
			{
                _ptr = pos;
                _token = tk;
				return null;
			}

			// check for empty Parameter list
			pos = _ptr;
			GetToken();
            if (_token.ID == TKID.CLOSE)
            {
                return null;
            }
			_ptr = pos;

			// get Parameters until we reach the end of the list
            var parms = new List<Expression>();
			var expr = ParseExpression();
			parms.Add(expr);
			while (_token.ID == TKID.COMMA)
			{
				expr = ParseExpression();
				parms.Add(expr);
			}

			// make sure the list was closed correctly
			if (_token.ID != TKID.CLOSE)
			{
                Throw();
			}

			// done
			return parms;
    	}
        Token GetMember()
        {
            // check whether next token is a MEMBER token ('.'), 
            // restore state and bail if it's not
            var pos = _ptr;
            var tk = _token;
            GetToken();
            if (_token.ID != TKID.PERIOD)
            {
                _ptr = pos;
                _token = tk;
                return null;
            }

            // skip member token
            GetToken();
            if (_token.Type != TKTYPE.IDENTIFIER)
            {
                Throw("Identifier expected");
            }
            return _token;
        }

		#endregion

		//---------------------------------------------------------------------------
		#region ** static helpers

        static void Throw()
        {
            Throw("Syntax error.");
        }
        static void Throw(string msg)
        {
            throw new Exception(msg);
        }

        #endregion
	}

    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    internal delegate object CalcEngineFunction(List<Expression> parms);
}
