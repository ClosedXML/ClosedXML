// Keep this file CodeMaid organised and cleaned
using System.Text;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Formula defined in a context-independent way. For example, is cell A1 contains formula '=B1'
    /// and A2 contains formula '=B2' they both share the same formula definition that may be represented
    /// as '=RC[1]', '=B1', '=B2', and so on.
    /// </summary>
    internal class XLFormulaDefinition
    {
        #region Public Constructors

        public XLFormulaDefinition(string formulaR1C1)
        {
            var res = _r1c1Parser.Parse(formulaR1C1, null);
            _formulaChunks = res.Item1;
            _references = res.Item2;
        }

        public XLFormulaDefinition(string formulaA1, XLAddress baseAddress)
        {
            var res = _a1Parser.Parse(formulaA1, baseAddress);
            _formulaChunks = res.Item1;
            _references = res.Item2;
        }

        #endregion Public Constructors

        #region Public Methods

        public string GetFormulaA1(IXLAddress baseAddress)
        {
            if (_formulaChunks.Length == 1)
                return _formulaChunks[0];

            var sb = new StringBuilder();
            for (int i = 0; i < _formulaChunks.Length - 1; i++)
            {
                sb.Append(_formulaChunks[i]);
                sb.Append(_references[i].ToStringA1(baseAddress));
            }

            sb.Append(_formulaChunks[_formulaChunks.Length - 1]);
            return sb.ToString();
        }

        public string GetFormulaR1C1()
        {
            //TODO Perhaps it is profitable to store R1C1 in a private field once it is built. Benchmark is needed
            if (_formulaChunks.Length == 1)
                return _formulaChunks[0];

            var sb = new StringBuilder();
            for (int i = 0; i < _formulaChunks.Length - 1; i++)
            {
                sb.Append(_formulaChunks[i]);
                sb.Append(_references[i].ToStringR1C1());
            }

            sb.Append(_formulaChunks[_formulaChunks.Length - 1]);
            return sb.ToString();
        }

        #endregion Public Methods

        #region Private Fields

        private static readonly XLFormulaDefinitionA1Parser _a1Parser = new XLFormulaDefinitionA1Parser();
        private static readonly XLFormulaDefinitionR1C1Parser _r1c1Parser = new XLFormulaDefinitionR1C1Parser();
        private readonly string[] _formulaChunks;
        private readonly IXLReference[] _references;

        #endregion Private Fields
    }
}
