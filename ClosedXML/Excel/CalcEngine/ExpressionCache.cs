#nullable disable

using System.Runtime.CompilerServices;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Caches expressions based on their string representation.
    /// This saves parsing time.
    /// </summary>
    /// <remarks>
    /// Uses weak references to avoid accumulating unused expressions.
    /// </remarks>
    internal sealed class ExpressionCache
    {
        private readonly ConditionalWeakTable<string, Formula> _cache;
        private readonly XLCalcEngine _ce;

        public ExpressionCache(XLCalcEngine ce)
        {
            _ce = ce;
            _cache = new ConditionalWeakTable<string, Formula>();
        }

        // gets the parsed version of a string expression
        public Formula this[string expression]
        {
            get
            {
                if (!_cache.TryGetValue(expression, out var formula))
                {
                    formula = _ce.Parse(expression);
                    _cache.Add(expression, formula);
                }
                return formula;
            }
        }
    }
}
